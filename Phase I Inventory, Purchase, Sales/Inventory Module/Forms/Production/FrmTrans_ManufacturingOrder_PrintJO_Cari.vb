Option Strict Off
Option Explicit On

Public Class FrmTrans_ManufacturingOrder_PrintJO_Cari
    Public IdentifierJO As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ScheduledOnly As Boolean = False

    Dim Obj As New ClsTrans_SPL
    Dim TBL As New DseTrans_ManufacturingOrder_PrintJO.tbltrans_cari_joDataTable
    Dim FindCriteria As ModGlobalVariable.EnumSearchCriteria = EnumSearchCriteria.ByNoDokumen
    Dim FindCriteriaJo As EnumSearchCriteriaJO = EnumSearchCriteriaJO.ByJoNumber

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.IdentifierJO = ""
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.IdentifierJO = Me.dgvHasil.SelectedRows(0).Cells(5).Value
            Else
                Me.IdentifierJO = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SPL_Cari_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Text = "Find Job Order"
            Me.cboCriteria.SelectedIndex = 0

            Me.txtCari.Focus()
            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If

            Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub txtCari_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.GotFocus
        txtCari.BackColor = BackColorAll
    End Sub

    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            If e.KeyValue = 13 And Me.txtCari.Text <> "" Then
                Me.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.FindCriteria, Me.FindCriteriaJo, Me.tslQuery, Me.ScheduledOnly, Me.dtpDari.Value, Me.dtpSampai.Value)
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                    Me.dgvHasil.Rows(0).Selected = True
                    '
                End If
            ElseIf e.KeyValue = 40 And Me.dgvHasil.RowCount > 0 Then
                Me.dgvHasil.Focus()
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.Rows.Item(0).Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Try
            Me.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, Me.FindCriteria, Me.FindCriteriaJo, Me.tslQuery, Me.ScheduledOnly, dtpDari.Value, Me.dtpSampai.Value)
            If Me.dgvHasil.Rows.Count > 0 Then
                Me.dgvHasil.ClearSelection()
                Me.dgvHasil.Rows(0).Selected = True
                '
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHasil.CellContentDoubleClick
        Try
            Me.Pilih()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvHasil.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.Pilih()
                Me.Close()
            ElseIf e.KeyValue = 38 Then
                If Me.dgvHasil.SelectedRows.Count > 0 And Me.dgvHasil.SelectedRows(0).Index = 0 Then
                    Me.txtCari.Focus()
                End If
            ElseIf e.KeyValue = 9 Then
                Me.txtCari.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.Pilih()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.LostFocus
        txtCari.BackColor = Me.txtPrimarySPL.BackColor
    End Sub

    Public Enum EnumSearchCriteriaJO As Integer
        ByJoNumber
    End Enum

    Private Sub cboCriteria_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCriteria.SelectedIndexChanged
        Try
            If Me.cboCriteria.Text = "JO Number" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteriaJO.ByJoNumber
            ElseIf Me.cboCriteria.Text = "Product Code" Then
                'Me.LblFrom.Visible = True
                'Me.lblTo.Visible = True
                'Me.dtpDari.Visible = True
                'Me.dtpSampai.Visible = True
                'Me.txtCari.Visible = False
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByProductCodeNo
            ElseIf Me.cboCriteria.Text = "Description" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByProductDesc
            ElseIf Me.cboCriteria.Text = "Remarks" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByRemarks
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpSampai_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpSampai.KeyDown
        If e.KeyValue = 13 Then
            Me.btnFind.PerformClick()
        End If
    End Sub

    Private Sub dgvHasil_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvHasil.MouseDoubleClick
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Pilih()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_RowHeaderMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvHasil.RowHeaderMouseDoubleClick
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Pilih()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub txtCari_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.GetJml(Me.txtCari.Text, Me.FindCriteria, Me.FindCriteriaJo, Me.ScheduledOnly, Me.dtpDari.Value, _
            Me.dtpSampai.Value) < 500 Then
                Me.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.FindCriteria, Me.FindCriteriaJo, Me.tslQuery, Me.ScheduledOnly, Me.dtpDari.Value, Me.dtpSampai.Value)
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                    Me.dgvHasil.Rows(0).Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
              EnumSearchCriteria, ByVal vSearchByJo As EnumSearchCriteriaJO, ByVal ScheduledOnly As Boolean, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
              Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long

        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_cari_joTableAdapter
        Dim RetVal As Nullable(Of Long)


        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteriaJO.ByJoNumber Then
                RetVal = DTA.ScalarQuery_GetJmlByJoNoLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByProductCodeNo Then
                RetVal = DTA.ScalarQuery_GetJmlByCodeNoLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByProductDesc Then
                RetVal = DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                RetVal = DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
   ByVal DGV As DataGridView, ByVal TBL As DseTrans_ManufacturingOrder_PrintJO.tbltrans_cari_joDataTable, _
   ByVal SearchCriteria As EnumSearchCriteria, ByVal SearchCriteriaJO As EnumSearchCriteriaJO, ByVal tsl As ToolStripLabel, ByVal ScheduledOnly As Boolean, _
   Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
   Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_cari_joTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If ScheduledOnly = False Then
                If SearchCriteria = EnumSearchCriteriaJO.ByJoNumber Then
                    QueryResult = DTA.FillBy_JONumberLike(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                    QueryResult = DTA.FillBy_ProductCodeLike(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                    QueryResult = DTA.FillBy_DescLike(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_RemarksLike(TBL, Criteria + "%")
                End If
            Else

            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView

            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).HeaderText = "Jo Number"
                DGV.Columns(1).HeaderText = "Product Code"
                DGV.Columns(2).HeaderText = "Description"

                DGV.Columns(4).HeaderText = "Remarks"
                DGV.Columns(5).HeaderText = "Identifier JO"
                DGV.Columns(5).Visible = False
                DGV.Columns(3).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > 10000000 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Ticks / TimeSpan.TicksPerSecond) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, 0, MidpointRounding.AwayFromZero)) + _
                    " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetDateMOByJoNo(ByVal vJoNo As String) As Date
        Dim DTA As New DseTrans_ManufacturingOrder_PrintJOTableAdapters.tbltrans_cari_joTableAdapter
        Dim Retval As New Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetDateMoByJONo(vJoNo)
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return "1/1/1950"
            End If
            Return "1/1/1950"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "1/1/1950"
        Finally
            DTA.Dispose()
        End Try
    End Function

End Class