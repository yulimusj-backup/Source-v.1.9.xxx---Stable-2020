Option Explicit On
Option Strict Off

Public Class frmLaporanVAT2025_InvalidNPWP
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public isEdit As Boolean = False
    Public _dt As New DataTable()


    Private Sub frmLaporanVAT2025_InvalidNPWP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ' Binding manual
            With dgvInvalidNPWP
                .Columns(Me.ColInvoiceNo.Index).DataPropertyName = "ColInvoiceNo"
                .Columns(Me.ColInvoiceDate.Index).DataPropertyName = "ColInvoiceDate"
                .Columns(Me.ColCustomerName.Index).DataPropertyName = "ColCustomerName"
                .Columns(Me.ColDocType.Index).DataPropertyName = "ColDocType"
                .Columns(Me.ColDocNo.Index).DataPropertyName = "ColDocNo"
                .Columns(Me.colPrimaryCustomer.Index).DataPropertyName = "colPrimaryCustomer"
                .Columns(Me.colPrimaryInvoice.Index).DataPropertyName = "colPrimaryInvoice"
            End With
            dgvInvalidNPWP.AutoGenerateColumns = False

            ' Ambil kolom ComboBox dari desain
            Dim comboCol As DataGridViewComboBoxColumn
            comboCol = CType(dgvInvalidNPWP.Columns(Me.ColDocType.Index), DataGridViewComboBoxColumn)
            ' Bersihkan item lama (kalau ada)
            comboCol.Items.Clear()
            ' Tambahkan 4 pilihan tetap
            comboCol.Items.Add("NPWP")
            comboCol.Items.Add("NIK")
            comboCol.Items.Add("PASSPORT")
            comboCol.Items.Add("OTHER ID")

            Me.dgvInvalidNPWP.AutoGenerateColumns = False
            Me.dgvInvalidNPWP.DataSource = Me._dt
            Me.dgvInvalidNPWP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvInvalidNPWP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvInvalidNPWP.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvInvalidNPWP_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvInvalidNPWP.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvInvalidNPWP_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvInvalidNPWP.CellContentClick
        Dim frmEditCustomer As New Core_Module.frmLaporanVAT2025_EditCustomer
        Try
            If e.ColumnIndex = Me.ColEdit.Index Then
                frmEditCustomer.UserCode = Me.UserCode
                frmEditCustomer.BackColorAll = Me.BackColorAll
                frmEditCustomer.PrimaryCustomer = Me.dgvInvalidNPWP.Rows(e.RowIndex).Cells(Me.colPrimaryCustomer.Index).Value
                frmEditCustomer.PrimaryInvoice = Me.dgvInvalidNPWP.Rows(e.RowIndex).Cells(Me.colPrimaryInvoice.Index).Value
                frmEditCustomer.TypeTransaction = Core_Module.frmLaporanVAT2025_EditCustomer.EnumTransactionType.TransactionSales
                frmEditCustomer.ShowDialog()
                Me.dgvInvalidNPWP.Rows(e.RowIndex).Cells(Me.ColDocType.Index).Value = frmEditCustomer.TypeDocument
                If frmEditCustomer.TypeDocument = "NPWP" Then
                    Me.dgvInvalidNPWP.Rows(e.RowIndex).Cells(Me.ColDocNo.Index).Value = frmEditCustomer.NPWP
                ElseIf frmEditCustomer.TypeDocument = "NIK" Then
                    Me.dgvInvalidNPWP.Rows(e.RowIndex).Cells(Me.ColDocNo.Index).Value = frmEditCustomer.NoNIKPASS
                ElseIf frmEditCustomer.TypeDocument = "PASSPORT" Then
                    Me.dgvInvalidNPWP.Rows(e.RowIndex).Cells(Me.ColDocNo.Index).Value = frmEditCustomer.NoNIKPASS
                ElseIf frmEditCustomer.TypeDocument = "OTHER ID" Then
                    Me.dgvInvalidNPWP.Rows(e.RowIndex).Cells(Me.ColDocNo.Index).Value = frmEditCustomer.NoNIKPASS
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As DataGridViewDataErrorEventArgs) _
    Handles dgvInvalidNPWP.DataError
        e.ThrowException = False
    End Sub

End Class