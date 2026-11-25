Option Explicit On
Option Strict Off

Public Class FormDummy
    Public vTblEnum As Integer = EnumTbl.shift
    Public vfrm As FrmTrans_SPL
    Public Enum EnumTbl As Integer
        shift
        setup
        detail
        approved
        approved_ot
    End Enum

    Private Sub FormDummy_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DataGridView1.AllowUserToDeleteRows = False
            Me.DataGridView1.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView1.DataSource = Nothing
            If Me.vTblEnum = EnumTbl.shift Then
                vfrm.TblEmpShift.DefaultView.RowFilter = Nothing
                Me.DataGridView1.DataSource = vfrm.TblEmpShift.DefaultView
            ElseIf Me.vTblEnum = EnumTbl.setup Then
                vfrm.TblSchSetup.DefaultView.RowFilter = Nothing
                Me.DataGridView1.DataSource = vfrm.TblSchSetup.DefaultView
            ElseIf Me.vTblEnum = EnumTbl.detail Then
                vfrm.TblSchDet.DefaultView.RowFilter = Nothing
                Me.DataGridView1.DataSource = vfrm.TblSchDet.DefaultView
            ElseIf Me.vTblEnum = EnumTbl.approved Then
                vfrm.TblApp.DefaultView.RowFilter = Nothing
                Me.DataGridView1.DataSource = vfrm.TblApp.DefaultView
            ElseIf Me.vTblEnum = EnumTbl.approved_ot Then
                'vfrm.TblAppOT.DefaultView.RowFilter = Nothing
                'Me.DataGridView1.DataSource = vfrm.TblAppOT.DefaultView
            End If
            Me.DataGridView1.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
            Me.DataGridView1.AllowUserToAddRows = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class