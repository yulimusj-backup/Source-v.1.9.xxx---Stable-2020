Option Explicit On
Option Strict Off

Partial Public Class FrmTools_MonthlyClosing
    Private Sub T1_TxtAccCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
         Handles T1_TxtAccCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.T1_TxtAccCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda(Me.T1_TxtAccCode.Text) = True Then
                        Me.T1_PrimaryAkunInv.Text = CStr(ObjAkun.GetPrimary(Me.T1_TxtAccCode.Text))
                        Me.T1_TxtAccCode.Text = ObjAkun.GetAccountCode(CDbl(Me.T1_PrimaryAkunInv.Text))
                        Me.T1_TxtAccCode.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.T1_btnFindAcc.PerformClick()
                    End If
                Else
                    Me.T1_btnFindAcc.PerformClick()
                End If
            End If
            Me.T1_TxtAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.T1_PrimaryAkunInv.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub T1_TxtAccCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles T1_TxtAccCode.Leave
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll

        Try
            If Me.T1_TxtAccCode.Text <> "" Then
                ''If Me.cboTipeTujuan.SelectedIndex = 0 Then
                If ObjAkun.ApakahSudahAda(Me.T1_TxtAccCode.Text) = True Then
                    Me.T1_PrimaryAkunInv.Text = CStr(ObjAkun.GetPrimary(Me.T1_TxtAccCode.Text))
                    Me.T1_TxtAccCode.Text = ObjAkun.GetAccountCode(CDbl(Me.T1_PrimaryAkunInv.Text))
                Else
                    If Me.T1_TxtAccCode.Text <> "" Then
                        ij.txtCari.Text = Me.T1_TxtAccCode.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.T1_PrimaryAkunInv.Text = CStr(ij.mPrimaryAkun)
                        Me.T1_TxtAccCode.Text = ObjAkun.GetAccountCode(CDbl(Me.T1_PrimaryAkunInv.Text))
                    Else
                        Me.T1_TxtAccCode.Clear()
                        Me.T1_PrimaryAkunInv.Text = "-1"
                        Me.T1_TxtAccCode.Focus()
                    End If
                End If
                ''End If
            Else
                Me.T1_TxtAccCode.Clear()
                Me.T1_PrimaryAkunInv.Text = "-1"
            End If
            Me.T1_TxtAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.T1_PrimaryAkunInv.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub T1_btnFindAcc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles T1_btnFindAcc.Click
        Dim ij As New AccModule.FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.T1_TxtAccCode.Text <> "" Then
                ij.txtCari.Text = Me.T1_TxtAccCode.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.T1_PrimaryAkunInv.Text = CStr(ij.mPrimaryAkun)
                Me.T1_TxtAccCode.Text = ObjAkun.GetAccountCode(CDbl(Me.T1_PrimaryAkunInv.Text))
                Me.T1_Btnfilter.Focus()
            Else
                Me.T1_TxtAccCode.Clear()
                Me.T1_PrimaryAkunInv.Text = "-1"
                Me.T1_TxtAccCode.Focus()
            End If
            Me.T1_TxtAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.T1_PrimaryAkunInv.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub T1_TxtAccCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles T1_TxtAccCode.GotFocus
        Try
            Me.T1_TxtAccCode.BackColor = Core_Module.GetSMESSColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub T1_TxtAccCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles T1_TxtAccCode.LostFocus
        Try
            Me.T1_TxtAccCode.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub T1_Btnfilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles T1_Btnfilter.Click
        Dim DTA As New DseToolsTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTools.DataTable1DataTable
        Dim row As DseTools.DataTable1Row
        Dim rowDGV As DataGridViewRow
        Dim DefPrec As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.T1_PrimaryAkunInv.Text = "-1" Or Me.T1_TxtAccCode.Text = "" Then
                MsgBox("Please specify a valid inventory account.", MsgBoxStyle.Information, "Information")
            End If
            DTA.FillBy(TBL, Me.ObjCD.GetLasDateOfTheMonth(Me.T1_dtpFrom.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.T1_dtpFrom.Value), Me.T1_PrimaryAkunInv.Text)

            Me.T1_dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.T1_dgv.Rows.Clear()
            DefPrec = ObjCurr.GetDecimal_DefaultCurrency
            For Each row In TBL.Rows
                Me.T1_dgv.Rows.Add()
                rowDGV = Me.T1_dgv.Rows(Me.T1_dgv.Rows.Count - 1)
                rowDGV.Cells(Me.ColT1_Amount.Index).Value = Me.ObjCD.DigitGrouping(row.At_Amount, DefPrec)
                rowDGV.Cells(Me.ColT1_AOPNo.Index).Value = row.No_Transaksi
                rowDGV.Cells(Me.ColT1_CodeNo.Index).Value = row.Product_Code
                rowDGV.Cells(Me.ColT1_Desc.Index).Value = row.Description
                rowDGV.Cells(Me.ColT1_No.Index).Value = Me.T1_dgv.Rows.Count
            Next
            Me.T1_dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            MsgBox("Data filtered", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub
End Class
