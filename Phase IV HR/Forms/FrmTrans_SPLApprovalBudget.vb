Option Explicit On
Option Strict Off

Imports Core_Module, System.Drawing, System.Windows.Forms, System.Threading
Public Class FrmTrans_SPLApprovalBudget
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Public PrimarySPL As Double = -1
    Public TblSplUsed As New DseTrans_SPL.ot_order_approvalDataTable
    Public rowSplUsed As DseTrans_SPL.ot_order_approvalRow
    Public TblSplAppDep2 As New DseTrans_SPL.ot_order_approval_2DataTable
    Public rowSplAppDep2 As DseTrans_SPL.ot_order_approval_2Row
    Public FromOvertimeOrder As Boolean = False
    Public ListPrimaryDept As New List(Of Double)
    Public PeriodFrom As Date = Now
    Public PeriodTo As Date = Now

    Private Sub FrmTrans_SPLApprovalBudget_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rowD As DataGridViewRow

        Try
            Me.DgvItem.AllowUserToResizeRows = False
            Me.ShowIcon = False

            Me.DgvItem.ReadOnly = True
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.DgvItem.Rows.Clear()
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If Me.FromOvertimeOrder = False Then
                'FromOvertimeOrder = false
                Me.TblSplAppDep2.DefaultView.RowFilter = "[Primary SPL] = " & Me.PrimarySPL & " AND [Apakah Approved] = 0"
                Me.TblSplAppDep2.DefaultView.Sort = "[Apakah Approved] DESC, [Department] ASC, [Period] ASC"
                For i As Integer = 0 To Me.TblSplAppDep2.DefaultView.Count - 1
                    rowSplAppDep2 = Me.TblSplAppDep2.DefaultView(i).Row
                    With rowSplAppDep2
                        Me.DgvItem.Rows.Add()
                        rowD = Me.DgvItem.Rows(Me.DgvItem.RowCount - 1)
                        rowD.Cells(Me.ColDeptPrimary.Index).Value = .Primary_Dept
                        rowD.Cells(Me.ColDept.Index).Value = .Department
                        rowD.Cells(Me.ColPeriod.Index).Value = Format(.Period, "MMM yyyy")
                        If .Status_Budget = FrmTrans_SPLApproval.EnumStatusBudget.OK Then
                            rowD.Cells(Me.ColStatus.Index).Value = "OK"
                        Else
                            rowD.Cells(Me.ColStatus.Index).Value = "Over"
                        End If
                    End With
                Next
            Else
                Dim ListPrimaryDepUrutan As New List(Of Double)
                Dim ParamPeriod As Date = "1/1/1950"
                For Each PrimaryDeptUrutan As Double In ListPrimaryDept
                    Me.TblSplAppDep2.DefaultView.RowFilter = "[Primary Dept] = " & PrimaryDeptUrutan
                    Me.TblSplAppDep2.DefaultView.Sort = "[Period] DESC, [Apakah Approved] DESC, [Department] ASC"
                    If Me.TblSplAppDep2.DefaultView.Count > 0 Then
                        For i As Integer = Me.TblSplAppDep2.DefaultView.Count - 1 To 0 Step -1
                            rowSplAppDep2 = Me.TblSplAppDep2.DefaultView(i).Row
                            With rowSplAppDep2
                                If .Period = Me.PeriodTo Or .Period = Me.PeriodFrom Then
                                    If ListPrimaryDepUrutan.Contains(PrimaryDeptUrutan) = False Or ParamPeriod <> .Period Then
                                        ListPrimaryDepUrutan.Add(PrimaryDeptUrutan)
                                        ParamPeriod = .Period

                                        Me.DgvItem.Rows.Add()
                                        rowD = Me.DgvItem.Rows(Me.DgvItem.RowCount - 1)
                                        rowD.Cells(Me.ColDeptPrimary.Index).Value = .Primary_Dept
                                        rowD.Cells(Me.ColDept.Index).Value = .Department
                                        rowD.Cells(Me.ColPeriod.Index).Value = Format(.Period, "MMM yyyy")
                                        If .Status_Budget = FrmTrans_SPLApproval.EnumStatusBudget.OK Then
                                            rowD.Cells(Me.ColStatus.Index).Value = "OK"
                                        Else
                                            rowD.Cells(Me.ColStatus.Index).Value = "Over"
                                        End If

                                        If .Period = Me.PeriodTo Then
                                            Exit For
                                        End If
                                    End If
                                End If
                            End With
                        Next
                    End If
                Next
            End If

            If Me.DgvItem.RowCount = 0 Then
                Me.TblSplAppDep2.DefaultView.RowFilter = "[Primary SPL] = " & Me.PrimarySPL & " AND [Apakah Approved] <> 0"
                Me.TblSplAppDep2.DefaultView.Sort = "[Apakah Approved] DESC, [Department] ASC, [Period] ASC"
                For Each dv As DataRowView In Me.TblSplAppDep2.DefaultView
                    rowSplAppDep2 = dv.Row
                    With rowSplAppDep2
                        Me.DgvItem.Rows.Add()
                        rowD = Me.DgvItem.Rows(Me.DgvItem.RowCount - 1)
                        rowD.Cells(Me.ColDeptPrimary.Index).Value = .Primary_Dept
                        rowD.Cells(Me.ColDept.Index).Value = .Department
                        rowD.Cells(Me.ColPeriod.Index).Value = Format(.Period, "MMM yyyy")
                        Me.TblSplUsed.DefaultView.RowFilter = "[Period] = #" & .Period.Date & "# AND[Primary Dept] = " & .Primary_Dept
                        rowD.Cells(Me.ColStatus.Index).Value = "OK"
                        If Me.TblSplUsed.DefaultView.Count > 0 Then
                            rowSplUsed = Me.TblSplUsed.DefaultView(0).Row
                            If rowSplUsed.Budget_Hitung < 0 Then
                                rowD.Cells(Me.ColStatus.Index).Value = "Over"
                            End If
                        End If
                    End With
                Next
            End If

            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            If Me.DgvItem.RowCount = 1 Then
                Me.StartPosition = FormStartPosition.Manual
                Me.Location = New Point(-3000, -3000)
                Me.Hide()
                LangsungTutupForm = True

                Dim arg = New DataGridViewCellEventArgs(Me.ColStatus.Index, 0)
                DgvItem_CellContentClick(Me.DgvItem, arg)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim LangsungTutupForm As Boolean = False
    Private Sub DgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellContentClick
        Dim i As New FrmTrans_SPLApprovalBudgetDetail

        Try
            If e.ColumnIndex = Me.ColStatus.Index Then
                i.PrimaryDept = CDbl(Me.DgvItem.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value)
                i.PrimarySPL = Me.PrimarySPL
                i.Period = CDate(Me.DgvItem.Item(Me.ColPeriod.Index, e.RowIndex).Value)
                i.TblSplUsed = Me.TblSplUsed
                i.TblSplAppDep2 = Me.TblSplAppDep2
                i.ListPrimaryDept = ListPrimaryDept
                i.Text &= " (" & Me.DgvItem.Item(Me.ColDept.Index, e.RowIndex).Value & " - " & Me.DgvItem.Item(Me.ColPeriod.Index, e.RowIndex).Value & ")"
                i.FromOvertimeOrder = Me.FromOvertimeOrder
                i.ShowDialog()

                If LangsungTutupForm = True Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub
End Class