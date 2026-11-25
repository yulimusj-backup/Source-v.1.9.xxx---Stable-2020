Option Explicit On
Option Strict Off
Public Class FrmTrans_SettingSalaryNPWP
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public typeNPWP As Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Company
    Public IsNPWPDigit As ClsTrans_SettingNPWPNo.EnumNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit
    Public TblItem As New DseTrans_SettingNPWP.TblErrorDataTable

    Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
    Dim ObjCD As New ClsCekData
    Dim Buffer As String = ""
    Dim Exiting As Boolean = False

    Public Enum EnumMsgReason
        NPWPdanIDCardTidakBolehKosong
        NPWP15Digit
        NPWP16Digit
        NPWPdanNIKHarussama
    End Enum

    Public Function GetMsgBoxReason(ByVal vMsgReason As FrmTrans_SettingSalaryNPWP.EnumMsgReason) As String
        Try
            If vMsgReason = EnumMsgReason.NPWPdanIDCardTidakBolehKosong Then
                Return "NPWP or NIK(Nomor Induk Kependudukan) cannot empty string"
            ElseIf vMsgReason = EnumMsgReason.NPWP16Digit Then
                Return "NPWP must consist of 16 digits"
            ElseIf vMsgReason = EnumMsgReason.NPWPdanNIKHarussama Then
                Return "NPWP and NIK must be the same"
            ElseIf vMsgReason = EnumMsgReason.NPWP15Digit Then
                Return "NPWP must consist of 15 digits"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub RefreshReason(ByVal rowIndex As Integer)
        Dim RowDGV As DataGridViewRow = Me.DgvItem.Rows(rowIndex)
        Dim Reason As String = ""
        Try
            With RowDGV
                Reason = ""
                If Reason = "" And String.IsNullOrEmpty(.Cells(Me.ColNPWP.Index).Value) = True Or .Cells(Me.ColNPWP.Index).Value = "-1" Or _
                String.IsNullOrEmpty(.Cells(Me.ColNIK.Index).Value) Or .Cells(Me.ColNIK.Index).Value = "-" Then
                    Reason = Me.GetMsgBoxReason(EnumMsgReason.NPWPdanIDCardTidakBolehKosong)
                End If

                If Me.IsNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit Then
                    If Reason = "" And ObjSetNPWP.ApakahSudah16Digit(.Cells(Me.ColNPWP.Index).Value, False) = False Then
                        Reason = Me.GetMsgBoxReason(EnumMsgReason.NPWP16Digit)
                    End If

                    If Reason = "" And .Cells(Me.ColNPWP.Index).Value <> .Cells(Me.ColNIK.Index).Value Then
                        Reason = Me.GetMsgBoxReason(EnumMsgReason.NPWPdanNIKHarussama)
                    End If
                Else
                    If Reason = "" And ObjSetNPWP.ApakahSudah15Digit(.Cells(Me.ColNPWP.Index).Value, False) = False Then
                        Reason = Me.GetMsgBoxReason(EnumMsgReason.NPWP15Digit)
                    End If
                End If

                .Cells(Me.ColReason.Index).Value = Reason
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function BlokSimbol(ByVal obj As TextBox, ByVal e As Windows.Forms.KeyPressEventArgs) As Boolean
        Try
            If e.KeyChar = "'" Or _
            e.KeyChar = "%" Or _
            e.KeyChar = "*" Or _
            e.KeyChar = "#" Then
                obj.ReadOnly = True
                obj.BackColor = Me.BackColorAll
            Else
                obj.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsikanDgv()
        Dim Row As DseTrans_SettingNPWP.TblErrorRow
        Dim rDGV As DataGridViewRow
        Try
            If TblItem.Rows.Count > 0 Then
                Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.DgvItem.Rows.Clear()

                For Each Row In TblItem.Rows
                    rDGV = Me.DgvItem.Rows(Me.DgvItem.Rows.Add)
                    With rDGV
                        .Cells(Me.ColKode.Index).Value = Row.Kode
                        .Cells(Me.ColNama.Index).Value = Row.Nama
                        .Cells(Me.ColNIK.Index).Value = Row.NIK
                        .Cells(Me.ColNPWP.Index).Value = Row.NPWP
                        .Cells(Me.ColNPWPLama.Index).Value = Row.NPWP
                        .Cells(Me.ColReason.Index).Value = Row.Reason
                        .Cells(Me.ColPrimaryKaryawan.Index).Value = Row.Primary_Karyawan
                    End With
                Next
                Me.ObjCD.SetNumber(Me.DgvItem, False)
            End If
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SalaryNPWPMassage_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Dialog As MsgBoxResult
        Try
            If Exiting = False Then
                Dialog = MsgBox("Exit?", MsgBoxStyle.OkCancel, "Confirmation")
                If Dialog = MsgBoxResult.Ok Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    Exit Sub
                End If
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SalaryNPWPMassage_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.IsikanDgv()

            If Me.TblItem.Rows.Count = 0 Then
                Me.Close()
            End If

            If Me.IsNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit Then
                Me.Label3.Text = "Based On Government Regulation No.: PMK 112/PMK.03/2022, NPWP No Contains of 16 Digit."
            Else
                Me.Label3.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColNPWP.Index Then 'Or e.ColumnIndex = Me.ColNIK.Index Then
                Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)

        Try
            With row
                If IsNothing(.Cells(Me.ColKode.Index).Value) = False Then

                    If e.ColumnIndex = Me.ColNPWP.Index Then
                        If Me.IsNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit Then
                            If Me.ObjSetNPWP.ApakahSudah16Digit(.Cells(Me.ColNPWP.Index).Value, False) = False Then
                                Me.GetMsgBoxReason(EnumMsgReason.NPWP16Digit)
                                .Cells(Me.ColNPWP.Index).Value = Me.Buffer
                                Exit Sub
                            End If
                        Else
                            If Me.ObjSetNPWP.ApakahSudah15Digit(.Cells(Me.ColNPWP.Index).Value, False) = False Then
                                Me.GetMsgBoxReason(EnumMsgReason.NPWP15Digit)
                                .Cells(Me.ColNPWP.Index).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If

                        .Cells(Me.ColNPWP.Index).Value = .Cells(Me.ColNPWP.Index).Value.ToString.Replace("%", "").Replace("*", "").Replace("'", "").Replace("#", "")
                        .Cells(Me.ColNIK.Index).Value = .Cells(Me.ColNPWP.Index).Value
                    End If

                    If e.ColumnIndex = Me.ColNIK.Index Then
                        .Cells(Me.ColNIK.Index).Value = .Cells(Me.ColNIK.Index).Value.ToString.Replace("%", "").Replace("*", "").Replace("'", "").Replace("#", "")
                    End If
                End If
                Me.RefreshReason(e.RowIndex)
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Dim Dialog As MsgBoxResult
        Try
            If Me.ApakahItemSudahBenar = True Then
                Dialog = MsgBox("Update this data?", MsgBoxStyle.YesNo, "Confirmation")
                If Dialog = MsgBoxResult.Yes Then

                    For Each row As DataGridViewRow In Me.DgvItem.Rows
                        With row
                            Me.ObjSetNPWP.UpdateNPWP_MasterKaryawan(.Cells(Me.ColNPWP.Index).Value, .Cells(Me.ColNPWPLama.Index).Value, ClsTrans_SettingNPWPNo.TipeNPWP.Individual, _
                             .Cells(Me.ColNIK.Index).Value, .Cells(Me.ColPrimaryKaryawan.Index).Value)
                        End With
                    Next

                    MsgBox("Successfully updated.", MsgBoxStyle.Information, "Information")
                    DialogResult = Windows.Forms.DialogResult.OK
                    Exiting = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahItemSudahBenar() As Boolean
        Dim Row As DataGridViewRow
        Try
            For Each Row In Me.DgvItem.Rows
                With Row
                    If String.IsNullOrEmpty(.Cells(Me.ColNIK.Index).Value) = True Or .Cells(Me.ColNIK.Index).Value = "-" Then
                        MsgBox(Me.GetMsgBoxReason(EnumMsgReason.NPWPdanIDCardTidakBolehKosong), MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    If String.IsNullOrEmpty(.Cells(Me.ColNPWP.Index).Value) = True Or .Cells(Me.ColNPWP.Index).Value = "-" Then
                        MsgBox(Me.GetMsgBoxReason(EnumMsgReason.NPWPdanIDCardTidakBolehKosong), MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    If Me.IsNPWPDigit = ClsTrans_SettingNPWPNo.EnumNPWPDigit._16Digit Then
                        If .Cells(Me.ColNIK.Index).Value <> .Cells(Me.ColNPWP.Index).Value Then
                            MsgBox(Me.GetMsgBoxReason(EnumMsgReason.NPWPdanNIKHarussama), MsgBoxStyle.Information, "Information")
                            Return False
                        End If

                        If Me.ObjSetNPWP.ApakahSudah16Digit(.Cells(Me.ColNPWP.Index).Value, False) = False Then
                            MsgBox(Me.GetMsgBoxReason(EnumMsgReason.NPWP16Digit), MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Else
                        If Me.ObjSetNPWP.ApakahSudah15Digit(.Cells(Me.ColNPWP.Index).Value, False) = False Then
                            MsgBox(Me.GetMsgBoxReason(EnumMsgReason.NPWP15Digit), MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End With
            Next

            If Me.DgvItem.Rows.Count = 0 Then
                MsgBox("At least one data item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Dim Dialog As MsgBoxResult
        Try
            Dialog = MsgBox("Exit?", MsgBoxStyle.OkCancel, "Confirmation")
            If Dialog = MsgBoxResult.Ok Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Exiting = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class