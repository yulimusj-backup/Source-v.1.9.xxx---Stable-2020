Option Explicit On
Option Strict Off

Public Class FrmMasterWorkUnit_AddDialog
    Public PrimaryEmployee As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FromAddDialog As Boolean = True
    Dim objWU As New ClsMasterWorkUnit
    Dim ObjCD As New Core_Module.ClsCekData
    Dim objEmployee As New ClsMasterKaryawan

    Private Sub FrmMasterWorkUnit_AddDialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.CboPilih.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOk.Click
        Dim hj As New FrmMasterWorkUnit
        Dim DTAItem As New DseMasterWorkUnitTableAdapters.tblmaster_workunit_item_employeeTableAdapter
        Dim PrimaryWorkShif As Double = -1
        Try
            hj.BackColorAll = Me.BackColorAll
            Me.Hide()
            
            If CboPilih.SelectedIndex = 0 Then
                hj.BringToFront()
                hj.Show()
                hj.btnNew.PerformClick()
                Dim rowDGV As DataGridViewRow
                If hj.TxtUnitDesc.Enabled = True Then
                    rowDGV = hj.DgvEmployee.Rows(hj.DgvEmployee.RowCount - 1)
                    With rowDGV
                        .Cells(hj.ColEmployee_No.Index).Value = rowDGV.Index + 1
                        .Cells(hj.ColEmployee_ShiftCode.Index).Value = ""
                        .Cells(hj.ColEmployee_EmployeeCode.Index).Value = objEmployee.GetCodeByPrimary(PrimaryEmployee)
                        .Cells(hj.ColEmployee_Name.Index).Value = objEmployee.GetFullNameByPrimary(PrimaryEmployee)
                        .Cells(hj.ColEmployee_Remarks.Index).Value = ""
                        .Cells(hj.ColEmployee_Primary.Index).Value = PrimaryEmployee
                        .Cells(hj.ColEmployee_PrimaryShift.Index).Value = "-1"
                    End With
                End If
            ElseIf CboPilih.SelectedIndex = 1 Then
                'hj.btnFind.PerformClick()
                Dim ij As New FrmMasterWorkUnit_Cari
                ij.BackColorAll = Me.BackColorAll
                ij.Status = "Show All"
                ij.ShowDialog()
                If Not ij.PrimaryWorkUnit = -1 Then
                    Dim nod As Nullable(Of Double)
                    nod = DTAItem.ScalarQuery_GetPrimaryWSByPrimaryMax(ij.PrimaryWorkUnit)
                    If nod.HasValue = False Then
                        nod = -1
                    End If
                    PrimaryWorkShif = nod.Value
                    If PrimaryWorkShif = -1 Then
                        Dim kj As New FrmMasterWorkShift_Cari(ClsMasterWorkShift.EnumTipeWorkshift.Production)
                        kj.BackColorAll = Me.BackColorAll
                        MsgBox("select work shift", MsgBoxStyle.Information, "Information")
                        kj.ShowDialog()
                        PrimaryWorkShif = kj.PrimaryWorkshift
                        kj = Nothing
                    End If
                    DTAItem.Insert(ij.PrimaryWorkUnit, PrimaryWorkShif, PrimaryEmployee, "")
                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                End If
                ij = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            'hj.Dispose()
        End Try
    End Sub
End Class