Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_ExplainRMReq_SetReqDate
    Public TBLAcuan As New DseTrans_ManufacturingOrder.DataTable7DataTable
    Public vPrimaryRM As Double = -1

    Dim rowAcuan As DseTrans_ManufacturingOrder.DataTable7Row
    Dim ObjMO1 As New ClsTrans_ManufacturingOrder
    Dim ObjSch As New ClsTrans_Schedule
    Dim LangsungClose As Boolean = False

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainRMReq_SetReqDate_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungClose = False Then
                If MsgBox("Close and discard data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    Me.LangsungClose = True
                Else
                    'langsung tutup - do nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_ExplainRMReq_SetReqDate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.cboJONumber.Items.Clear()
            Me.TBLAcuan.DefaultView.Sort = "[JO No]"
            For Each DV As Data.DataRowView In Me.TBLAcuan.DefaultView
                rowAcuan = DV.Row
                Me.cboJONumber.Items.Add(rowAcuan.JO_No)
            Next
            Me.cboJONumber.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.ApakahDataBenar() = True Then
                If MsgBox("Set raw material required date to: " + Format(Me.dtpRequiredDate.Value, "dd MMM yyyy") + "?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.LangsungClose = True
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim Dummy As String = ""
        Dim PrimaryMO As Double = -1
        Dim TglMO, TglJadwal As Date

        Try
            For Each rowAcuan In Me.TBLAcuan.Rows
                If Me.cboJONumber.Text = rowAcuan.JO_No Then
                    Dummy = rowAcuan.Identifier_JO
                    PrimaryMO = Dummy.Substring(0, Dummy.IndexOf("-"))
                    TglMO = Me.ObjMO1.GetDate(PrimaryMO).Date
                    TglJadwal = Me.ObjSch.GetStartTime1(rowAcuan.Identifier_JO, Me.vPrimaryRM).Date
                    If TglJadwal <> #1/1/1950# Then
                        If TglJadwal <= TglMO Then
                            TglMO = TglJadwal - TimeSpan.FromDays(1) ' minus 1 hari, asumsi rm harus tersedia 1 hari sebelum produksi dimulai
                        End If
                    End If

                    'jika sudah terjadwal, pastikan tanggal lebih kecil dari tanggal mulai jadwal
                    If Me.dtpRequiredDate.Value.Date < TglMO Then
                        MsgBox("Please specify a date value with minimum value: " + Format(TglMO, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    'sampe sini ok
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
End Class