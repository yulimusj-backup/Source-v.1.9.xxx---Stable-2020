Option Explicit On
Option Strict Off

Public Class FrmTrans_SalesDashboard
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim AllowToSeeAllUser As Boolean = True

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCRM_Prospect.Click
        Dim i As New FrmMasterProspect
        i.ShowDialog()
        i.Dispose()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim i As New FrmMasterProspect_Opportunity
        i.ShowDialog()
        i.Dispose()
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCRM_Quotation.Click

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCRM_Opportunity.Click
        Dim i As New FrmTrans_Opportunity
        i.ShowDialog()
        i.Dispose()
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvProspectActivity.DataError

        Try
            'berkaitan pencegahan data error kalo loading dan menentukan tipe opportunity
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvProspectActivity.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvProspectActivity.Rows(Me.dgvProspectActivity.Rows.Count - 1)
            rowDGV.Cells(Me.ColActivity_Type.Index).Value = "Prospect"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_SalesDashboard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            UpdateKolomSF()
            If Me.txtCodeNo.Text <> "" Then
                If Me.ObjProduct.ApakahSudahAda(Me.txtCodeNo.Text) = True Then
                    Me.txtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.SetGBO()
                End If
            End If
            Me.dgvDOItem.ReadOnly = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkIncUnAuthPO_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkIncUnAuthPO.CheckedChanged
        Dim mDecimal As Integer = 0
        Try
            If Me.txtPrimaryCodeNo.Text <> "-1" Then
                If Me.ChkIncUnAuthPO.Checked = True Then
                    Me.UnAuthPO = Inventory_Module.ClsTransaksiPOItem.IncludeUnAuthPO.IncludeUnAuthPO
                    If Me.dgvItemPOB.RowCount > 0 Then
                        Me.dgvItemPOB.Columns(13).Visible = True
                    End If
                Else
                    Me.UnAuthPO = Inventory_Module.ClsTransaksiPOItem.IncludeUnAuthPO.ExcludeUnAuthPO
                    If Me.dgvItemPOB.RowCount > 0 Then
                        Me.dgvItemPOB.Columns(13).Visible = False
                    End If
                End If

                mDecimal = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
                Me.ObjPO.IsiDGVPOB(Me.txtPrimaryCodeNo.Text, Me.dgvItemPOB, Me.TBL_POB, Me.UnAuthPO)
                Me.TxtPOB.Text = Me.ObjCd.DigitGrouping(Me.ObjPO.GetOrderBalance(Me.txtPrimaryCodeNo.Text, Me.UnAuthPO), mDecimal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabPage4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage4.Click

    End Sub
End Class