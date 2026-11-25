Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_NewSch6
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TblRef As New DseTrans_Schedule2.TblRefDataTable
    Dim rowRef As DseTrans_Schedule2.TblRefRow
    Dim ObjCD As New Core_Module.ClsCekData
    Dim LansungClose As Boolean = False
    Dim Buffer As String = ""
    Dim TBLRef2 As New DseTrans_Schedule2.TblRef2DataTable
    Dim rowRef2 As DseTrans_Schedule2.TblRef2Row
    Dim DefPrec As Integer = 2

    Private Sub FrmTrans_ProductionSchedule_NewSch6_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ListPrimary As New List(Of Double)

        Try
            'isikan tbl ref 2
            For Each rowRef In TblRef.Rows
                If ListPrimary.Contains(rowRef.Primary_Process) = False Then
                    ListPrimary.Add(rowRef.Primary_Process)

                    rowRef2 = TBLRef2.NewRow
                    With rowRef2
                        .Primary_Process = rowRef.Primary_Process
                        .Process_Code = rowRef.Process_Code
                    End With
                    TBLRef2.Rows.Add(rowRef2)
                End If
            Next

            Me.cboProcess.Items.Clear()
            For Each rowRef In TblRef.Rows
                Me.cboProcess.Items.Add(rowRef.Process_Code)
            Next
            'Me.cboProcess.SelectedIndex = 0
            Me.dgvWU.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            'tampilkan ke cbo
            Me.cboProcess.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary.Clear()
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch6_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LansungClose = False Then
                If MsgBox("Cancel data change?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Else
                    e.Cancel = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If MsgBox("Accept data change?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.LansungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboProcess_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProcess.SelectedValueChanged
        Dim rowDGV As DataGridViewRow

        Try
            If Me.cboProcess.SelectedIndex = -1 Then
                Me.dgvWU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvWU.Rows.Clear()
                Me.dgvWU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Else
                Me.dgvWU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvWU.Rows.Clear()

                Me.TBLRef2.DefaultView.RowFilter = "[Process Code] = '" + Me.cboProcess.Text + "'"
                If Me.TBLRef2.DefaultView.Count > 0 Then
                    rowRef2 = TBLRef2.DefaultView(0).Row

                    Me.TblRef.DefaultView.RowFilter = "[Primary Process] =" + Me.rowRef2.Primary_Process.ToString
                    Me.TblRef.DefaultView.Sort = "[Work Unit Code]"

                    For Each DV As Data.DataRowView In TblRef.DefaultView
                        rowRef = DV.Row

                        Me.dgvWU.Rows.Add()
                        rowDGV = Me.dgvWU.Rows(Me.dgvWU.Rows.Count - 1)
                        rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvWU.Rows.Count
                        rowDGV.Cells(Me.ColPrimaryWorkUnit.Index).Value = rowRef.Primary_WU
                        rowDGV.Cells(Me.ColRate.Index).Value = Me.ObjCD.DigitGrouping(rowRef.Rate, DefPrec)
                        rowDGV.Cells(Me.ColWorkUnitCode.Index).Value = rowRef.Work_Unit_Code
                        rowDGV.Cells(Me.ColWorkUnitDesc.Index).Value = rowRef.Work_Unit_Desc
                    Next
                End If

                Me.dgvWU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWU_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWU.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColRate.Index Then
                Me.Buffer = Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWU_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWU.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvWU.Rows(e.RowIndex)

        Try
            With row
                If IsNothing(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    
                    If e.ColumnIndex = Me.ColRate.Index Then
                        'cek bilang positif
                        If Me.ObjCD.CekBilanganPositif(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please specify a positive number.", MsgBoxStyle.Information, "Information")
                            Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        'grouping
                        Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value, DefPrec)

                        If Me.ObjCD.CekBilanganPositif(Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please specify a positive number.", MsgBoxStyle.Information, "Information")
                            Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                Else
                    Me.dgvWU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class