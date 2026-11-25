Public Class FrmVersionControl 
    Dim ObjVersionControl As New Core_Module.Clstrans_VersionControl
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmVersionControl_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rowDGV As DataGridViewRow
        Dim DTA As New Core_Module.DseTrans_Version_ControlTableAdapters.tbltrans_version_controlTableAdapter
        Dim TBL As New Core_Module.DseTrans_Version_Control.tbltrans_version_controlDataTable
        Dim row As Core_Module.DseTrans_Version_Control.tbltrans_version_controlRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.Rows.Clear()

            DTA.Fill(TBL)
            For Each row In TBL.Rows
                row.expiry = Me.ObjVersionControl.GetExpiry(row.license_expiry, row.product_key)

                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColComputerName.Index).Value = row.computer_name
                rowDGV.Cells(Me.ColExpiry.Index).Value = row.expiry
                rowDGV.Cells(Me.ColLastUserLogin.Index).Value = row.last_user_login
                rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(Me.ColProductKey.Index).Value = row.product_key
                rowDGV.Cells(Me.ColLastLoginTime.Index).Value = row.last_login_timestamp
                rowDGV.Cells(Me.ColVersion.Index).Value = row.version
            Next

            Me.ColExpiry.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.ColLastLoginTime.DefaultCellStyle.Format = "dd MMM yyyy HH.mm"

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class