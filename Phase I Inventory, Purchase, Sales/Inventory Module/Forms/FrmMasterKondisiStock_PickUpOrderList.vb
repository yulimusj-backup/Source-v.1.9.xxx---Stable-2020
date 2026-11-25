Option Explicit On
Option Strict Off

Public Class FrmMasterKondisiStock_PickUpOrderList
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Public VPrimaryPR As Double = -1
    Public vPrimaryCodeNo As Double = -1
    Public vPrimaryPrItemDel As Double = -1
    Dim ObjPR As New ClsTransaksiPR
    Public vTblPuo2 As DseTransaksiPR.Tbl_PickUpOrder2DataTable
    
    Private Sub FrmMasterKondisiStock_PickUpOrderList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ObjPR.IsiDGVPickUpList(Me.DgvDetail, VPrimaryPR, vPrimaryPrItemDel, Me, vTblPuo2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvDetail_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvDetail.CellContentClick
        Dim vPrimary As Double = -1

        Try
            If e.ColumnIndex = -1 Or e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.Col_PickUpOrderNo.Index Then
                vPrimary = Me.DgvDetail.Item(Me.Col_PrimaryPickUpOrder.Index, e.RowIndex).Value
                Dim ij As FrmTransaksiPR_PickUpOrder2
                Dim j As Windows.Forms.Form

                For Each j In Me.ParentForm.MdiChildren
                    ij = TryCast(j, FrmTransaksiPR_PickUpOrder2)
                    If IsNothing(ij) = False Then
                        ij.BringToFront()
                        ij.Focus()
                        ij.ShowTransaction(vPrimary)
                        Me.Close()
                        Exit Sub
                    End If
                Next

                Dim i As New FrmTransaksiPR_PickUpOrder2
                ' Increment the caption counter.
                i.UserCode = Me.UserCode.ToString
                ' Set the caption to be unique.
                i.BAckColorAll = Core_Module.GetSMESSBackColorAll
                i.MdiParent = Me.ParentForm
                i.Show()
                i.ShowTransaction(vPrimary)
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub
End Class