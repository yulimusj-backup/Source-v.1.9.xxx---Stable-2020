Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms, System.Threading
Public Class FrmTrans_SPLApprovalBudgetDetail
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Public PrimaryDept As Double = -1
    Public PrimarySPL As Double = -1
    Public Period As Date = Now
    Public TblSplUsed As New DseTrans_SPL.ot_order_approvalDataTable
    Public rowSplUsed As DseTrans_SPL.ot_order_approvalRow
    Public TblSplAppDep2 As New DseTrans_SPL.ot_order_approval_2DataTable
    Public rowSplAppDep2 As DseTrans_SPL.ot_order_approval_2Row
    Public FromOvertimeOrder As Boolean = False
    Public ListPrimaryDept As New List(Of Double)

    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmTrans_SPLApprovalBudgetDetail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rowD As DataGridViewRow
        Dim TotalUsed As Double = 0

        Try
            Me.DgvItem.AllowUserToResizeRows = False
            Me.ShowIcon = False
            Me.TblSplUsed.DefaultView.RowFilter = "[Period] = #" & Period.Date & "# AND[Primary Dept] = " & Me.PrimaryDept
            If Me.TblSplUsed.DefaultView.Count > 0 Then
                rowSplUsed = Me.TblSplUsed.DefaultView(0).Row
                Me.TxtBudget.Text = Me.ObjCD.Convert_ToString(rowSplUsed.Budget)
                'Me.TxtUsed.Text = Me.ObjCD.Convert_ToString(rowSplUsed.Used)
            End If

            Me.DgvItem.ReadOnly = True
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.DgvItem.Rows.Clear()
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.TblSplAppDep2.DefaultView.RowFilter = "[Primary Dept] = " & Me.PrimaryDept & " AND [Period] = #" & Me.Period & "#" ' AND [Apakah Approved] = 0"
            Me.TblSplAppDep2.DefaultView.Sort = "[Apakah Approved] DESC, [Spl Date] ASC, [Urutan] ASC"
            For Each dv As DataRowView In Me.TblSplAppDep2.DefaultView
                rowSplAppDep2 = dv.Row
                With rowSplAppDep2
                    If .Apakah_Approved = 2 And .Proposed = 0 And .Used = 0 Then
                        GoTo SkipRows
                    End If

                    Me.DgvItem.Rows.Add()
                    rowD = Me.DgvItem.Rows(Me.DgvItem.RowCount - 1)
                    rowD.Cells(Me.ColSpl.Index).Value = .Spl_No
                    rowD.Cells(Me.ColDate.Index).Value = Format(.Spl_Date, "dd MMM yyyy")
                    rowD.Cells(Me.ColProposed.Index).Value = Me.ObjCD.Convert_ToString(.Proposed)
                    TotalUsed += .Used
                    rowD.Cells(Me.ColUsed.Index).Value = Me.ObjCD.Convert_ToString(.Used)
                    If .Balance >= 0 Then
                        rowD.Cells(Me.ColBalance.Index).Value = Me.ObjCD.Convert_ToString(.Balance)
                    Else
                        rowD.Cells(Me.ColBalance.Index).Value = "-" & Me.ObjCD.Convert_ToString(.Balance * -1)
                    End If
                    If rowSplAppDep2.Apakah_Approved = 1 Then
                        rowD.Cells(Me.ColStatus.Index).Value = "Approved"
                    ElseIf rowSplAppDep2.Apakah_Approved = 0 Then
                        rowD.Cells(Me.ColStatus.Index).Value = "Waiting For Approval"
                    Else
                        rowD.Cells(Me.ColStatus.Index).Value = "Use OT without overtime order"
                    End If
                    rowD.Cells(Me.ColPrimary.Index).Value = .Primary_Spl

                    If FromOvertimeOrder = False Then
                        If .Primary_Spl = Me.PrimarySPL Then
                            Exit For
                        End If
                    End If

SkipRows:
                End With
            Next
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.TxtUsed.Text = Me.ObjCD.Convert_ToString(TotalUsed)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUsed_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtUsed.MouseClick
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUsed_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUsed.MouseHover
        Try
            Me.Cursor = Cursors.Hand
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUsed_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUsed.MouseLeave
        Try
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub
End Class