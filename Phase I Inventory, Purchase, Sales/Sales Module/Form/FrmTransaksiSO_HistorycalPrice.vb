Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiSO_HistorycalPrice
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public mPrimaryProduct As Double = -1

    Dim ObjCustomer As New clsMasterCustomer
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim TBL As New DseTransaksiSO.DataTable6DataTable

    Public Sub IsiDGV()
        Try
            Me.ObjSO.IsiDGV_HistoricalPrice(Me.txtPrimaryCustomer.Text, Me.mPrimaryProduct, Me.dgvItem, Me.TBL, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus
        Try
            If Me.txtCustomerCode.ReadOnly = False Then
                Me.txtCustomerCode.BackColor = Me.BackColorAll
            Else
                Me.txtCustomerCode.BackColor = Me.TxtCustomerName.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus
        Try
            If Me.txtCustomerCode.ReadOnly = False Then
                Me.txtCustomerCode.BackColor = Me.txtPrimaryCustomer.BackColor
            Else
                Me.txtCustomerCode.BackColor = Me.TxtCustomerName.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.dgvItem.Focus()
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.dgvItem.Focus()
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtCustomerCode.Focus()
                    End If
                End If
                Me.IsiDGV()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCustomerCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) Then
                    Dim PrimaryBaru As Decimal = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    If PrimaryBaru <> Me.txtPrimaryCustomer.Text Then
                        Me.txtPrimaryCustomer.Text = PrimaryBaru
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.dgvItem.Focus()
                    End If
                Else
                    If Me.txtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.dgvItem.Focus()
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtCustomerCode.Focus()
                    End If
                End If
            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.txtPrimaryCustomer.Text = "-1"
            End If
            Me.IsiDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomer.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCustomerCode.Text <> "" Then
                i.txtCari.Text = Me.txtCustomerCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.dgvItem.Focus()
            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtCustomerCode.Focus()
            End If
            Me.IsiDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub FrmTransaksiSO_HistorycalPrice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Me.ShowIcon = False
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            For Each Col As DataGridViewColumn In Me.dgvItem.Columns
                Col.SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            Me.IsiDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class