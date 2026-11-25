Imports System.Windows.Forms

Public Class FrmTrans_StockBundling_Quick2
    Public IdentifierWO As String = ""
    Public IdentifierJO As String = ""
    Public TBLWU As New DseTrans_ManufacturingOrder.DataTable10DataTable

    Dim DTAWU As New DseTrans_ManufacturingOrderTableAdapters.DataTable10TableAdapter
    
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Try
            If MsgBox("Accept work unit and close form?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub FrmTrans_StockBundling_Quick2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DTAWU.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAWU.Fill(Me.TBLWU, Me.IdentifierJO, Me.IdentifierWO)

            Me.cboWorkUnit.DataSource = Me.TBLWU.DefaultView
            Me.cboWorkUnit.SelectedIndex = 0
            Me.cboWorkUnit.DisplayMember = "Work Unit"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
