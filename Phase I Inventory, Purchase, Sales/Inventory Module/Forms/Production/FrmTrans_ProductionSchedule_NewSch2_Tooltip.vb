''ok-rosadin-r5b
Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_NewSch2_Tooltip
    Public tooltipColor As Color = Color.PeachPuff
    Public job_order_number As String = ""
    Public wo_no As String = ""
    Public process_code As String = ""
    Public work_unit_code As String = ""
    Public code_no As String = ""
    Public description As String = ""
    Public rate_unit_per_hour As Decimal = 0
    Public qty As Decimal = 0
    Public unit_code As String = ""
    Public work_hour As Decimal = 0
    Public start As Date = #1/1/1950#
    Public finish As Date = #1/1/1950#
    Public progress As Decimal = 0
    Public line_code As String = ""
    Dim decimal_ As Integer = 2
    Dim objCd As New Core_Module.ClsCekData

    Private Sub FrmTrans_ProductionSchedule_NewSch2_Tooltip_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim r As DataGridViewRow
        Dim dgvHeight As Integer = 0
        Try
            Me.DataGridView1.Anchor = AnchorStyles.Top And AnchorStyles.Left
            Me.DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.None
            If job_order_number <> "" Then
                dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            End If
            If wo_no <> "" Then
                dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            End If
            If process_code <> "" Then
                dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            End If
            If work_unit_code <> "" Then
                dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            End If
            If code_no <> "" Then
                dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            End If
            If description <> "" Then
                dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            End If
            If rate_unit_per_hour <> 0 Then
                dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            End If
            If qty <> 0 Then
                dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            End If
            'If unit_code <> "" Then
            '   dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            'End If
            If work_hour <> 0 Then
                dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            End If
            If start <> #1/1/1950# Then
                dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            End If
            If finish <> #1/1/1950# Then
                dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            End If
            'If progress <> 0 Then
            dgvHeight = dgvHeight + Me.DataGridView1.RowTemplate.Height
            'End If
            Me.Height = dgvHeight
            Me.DataGridView1.Height = dgvHeight
            Me.DataGridView1.Location = New Point(0, 0)

            If job_order_number <> "" Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "JO No"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = job_order_number
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            If wo_no <> "" Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "WO No"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = wo_no
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            If process_code <> "" Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "Process Code"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = process_code
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            If work_unit_code <> "" Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "Work Unit"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = work_unit_code
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            If code_no <> "" Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "Code No"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = code_no
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            If description <> "" Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "Description"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = description
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            If rate_unit_per_hour <> 0 Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "Rate (Unit/Hour)"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = Me.objCd.DigitGrouping(rate_unit_per_hour, decimal_)
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            If qty <> 0 Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "Produce Qty"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = Me.objCd.DigitGrouping(qty, decimal_) & " " & unit_code
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            'If unit_code <> "" Then
            '    DataGridView1.Rows.Add()
            '    r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
            '    r.Cells(Column1.Index).Value = "Unit Code"
            '    r.Cells(Column1.Index).Style.BackColor = tooltipColor
            '    r.Cells(Column2.Index).Value = ":"
            '    r.Cells(Column2.Index).Style.BackColor = tooltipColor
            '    r.Cells(Column3.Index).Value = unit_code
            '    r.Cells(Column3.Index).Style.BackColor = tooltipColor
            'End If
            If work_hour <> 0 Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "Work Hour"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = Me.objCd.Convert_ToString(work_hour)
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            If start <> #1/1/1950# Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "Start"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = Format(start, "dd MMM yyyy HH:mm")
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            If finish <> #1/1/1950# Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "Finish"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = Format(finish, "dd MMM yyyy HH:mm")
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            If progress > -1 Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "Progress"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = Me.objCd.DigitGrouping(progress, 2) & " %"
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
            If line_code <> "" Then
                DataGridView1.Rows.Add()
                r = DataGridView1.Rows(DataGridView1.Rows.Count - 1)
                r.Cells(Column1.Index).Value = "Line Code"
                r.Cells(Column1.Index).Style.BackColor = tooltipColor
                r.Cells(Column2.Index).Value = ":"
                r.Cells(Column2.Index).Style.BackColor = tooltipColor
                r.Cells(Column3.Index).Value = line_code
                r.Cells(Column3.Index).Style.BackColor = tooltipColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        e.Cancel = True
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch2_Tooltip_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LostFocus
        Me.Close()
    End Sub

End Class