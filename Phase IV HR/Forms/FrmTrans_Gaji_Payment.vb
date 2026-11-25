Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing, Core_Module
Public Class FrmTrans_Gaji_Payment

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Def_Prec As Double = 0
    Public PrimarySalary As Double
    Public Baru As Boolean
    Public NonAktif As Boolean = False
    Public FrmSalary As FrmTrans_Gaji
    Dim Buffer As String = ""
    Dim EraseRowIndex_PayAccount As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjAkun As New AccModule.ClsMasterAkun
    Dim Obj As New ClsTrans_Payment
    Dim ObjItem As New ClsTrans_Payment_Item

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadData()
        Try
            Me.DGVPayment.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.NonAktif = False Then
                Me.DGVPayment.ReadOnly = True
            Else
                Me.DGVPayment.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_Payment_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Amount As Decimal = 0
        Dim RowDGV As DataGridViewRow
        Try
            Me.DGVPayment.AllowUserToAddRows = False
            Me.FrmSalary.TxtPayment.Text = "0"
            If Me.DGVPayment.Rows.Count > 0 Then
                For Each RowDGV In DGVPayment.Rows
                    Amount += CDec(RowDGV.Cells(Me.ColDetailPayment_Amount.Index).Value)
                Next
                Me.FrmSalary.TxtPayment.Text = Me.ObjCD.DigitGrouping(Amount, Me.Def_Prec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_OtherDeduction_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.LoadData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVPayment_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVPayment.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColDetailPayment_AccountCode.Index _
            Or e.ColumnIndex = Me.ColDetailPayment_Amount.Index _
            Or e.ColumnIndex = Me.ColDetailPayment_Remarks.Index Then
                Me.Buffer = Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemInput_PayAccount(ByVal vPrimaryAkun As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.DGVPayment.Rows
                If rowDGV.Cells(Me.ColDetailPayment_PrimaryAkun.Index).Value = vPrimaryAkun Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub DGVPayment_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVPayment.CellContentClick
        Dim RowDGV As DataGridViewRow
        Dim Hasil As Decimal = 0
        Try
            If Me.NonAktif = True Then
                If IsNothing(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColTambahan.Index Then
                        If Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value.ToString = "" Or _
                        Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value.ToString = "-1" Then
                            MsgBox("Please input account code.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                        Hasil = 0
                        For i As Integer = 0 To Me.DGVPayment.Rows.Count - 2
                            RowDGV = Me.DGVPayment.Rows(i)
                            If e.RowIndex <> RowDGV.Index Then
                                Hasil += RowDGV.Cells(Me.ColDetailPayment_Amount.Index).Value
                            End If
                        Next
                        Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.FrmSalary.TxtSalaryReceived.Text) - Hasil, Me.Def_Prec)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVPayment_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVPayment.CellEndEdit
        Dim Row As DataGridViewRow = Me.DGVPayment.Rows(e.RowIndex)
        Dim Index As Integer
        Dim IdentifierStr As String
        Dim BufferPrimaryAkunPay As Double = -1
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    ''Cell kode 
                    If e.ColumnIndex = Me.ColDetailPayment_AccountCode.Index Then
                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Index = Me.GetIndexItemInput_PayAccount(Me.ObjAkun.GetPrimary(Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value))
                            If Index < 0 Then
                                'Batasan Akun
                                If Me.Obj.ApakahAkunPaymentCashNBankAtoHutang(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value) = False Then
                                    Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_PayAccount = e.RowIndex
                                        Me.BackgroundWorker_Payment.RunWorkerAsync()
                                    Else
                                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DGVPayment)
                                    Exit Sub
                                End If
                                BufferPrimaryAkunPay = CDbl(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value)
                                Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.DGVPayment.Item(Me.ColDetailPayment_AccountDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value)
                                If String.IsNullOrEmpty(Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value) = True Then
                                    Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("1", Me.Def_Prec)
                                End If
                                IdentifierStr = ObjItem.GetIdentifierUrutan_BuatDGVCellEndEdit(Me.DGVPayment, Me.ColDetailPayment_Identifier.Index, Me.Baru, Me.PrimarySalary, Row.Index)
                                Me.DGVPayment.Item(Me.ColDetailPayment_Identifier.Index, e.RowIndex).Value = IdentifierStr
                                Me.DGVPayment.Item(Me.ColDetailPayment_IndexRowSalary.Index, e.RowIndex).Value = Row.Index
                            Else
                                If e.RowIndex <> Index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_PayAccount = e.RowIndex
                                        Me.BackgroundWorker_Payment.RunWorkerAsync()
                                    Else
                                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DGVPayment)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New AccModule.FrmMasterAkun_Cari
                            ij.txtCari.Text = Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.mPrimaryAkun <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItemInput_PayAccount(ij.mPrimaryAkun)
                                If Index < 0 Then
                                    'Batasan Akun
                                    If Me.Obj.ApakahAkunPaymentCashNBankAtoHutang(ij.mPrimaryAkun) = False Then
                                        Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = ""
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_PayAccount = e.RowIndex
                                            Me.BackgroundWorker_Payment.RunWorkerAsync()
                                        Else
                                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DGVPayment)
                                        Exit Sub
                                    End If
                                    Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                                    Me.DGVPayment.Item(Me.ColDetailPayment_AccountDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                                    If String.IsNullOrEmpty(Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value) = True Then
                                        Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("1", Me.Def_Prec)
                                    End If
                                    IdentifierStr = ObjItem.GetIdentifierUrutan_BuatDGVCellEndEdit(Me.DGVPayment, Me.ColDetailPayment_Identifier.Index, Me.Baru, Me.PrimarySalary, Row.Index)
                                    Me.DGVPayment.Item(Me.ColDetailPayment_Identifier.Index, e.RowIndex).Value = IdentifierStr
                                    Me.DGVPayment.Item(Me.ColDetailPayment_IndexRowSalary.Index, e.RowIndex).Value = Row.Index
                                    BufferPrimaryAkunPay = CDbl(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value)
                                    Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                                Else
                                    If e.RowIndex <> Index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_PayAccount = e.RowIndex
                                            Me.BackgroundWorker_Payment.RunWorkerAsync()
                                        Else
                                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.DGVPayment)
                                        Exit Sub
                                    Else
                                        Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                                        Me.DGVPayment.Item(Me.ColDetailPayment_AccountDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                                        'Batasan Akun
                                        If Me.Obj.ApakahAkunPaymentCashNBankAtoHutang(ij.mPrimaryAkun) = False Then
                                            Me.DGVPayment.Item(Me.ColDetailPayment_AccountCode.Index, e.RowIndex).Value = ""
                                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                                Me.EraseRowIndex_PayAccount = e.RowIndex
                                                Me.BackgroundWorker_Payment.RunWorkerAsync()
                                            Else
                                                Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            End If
                                            Me.ObjCD.SetNumber(Me.DGVPayment)
                                            Exit Sub
                                        End If
                                        If String.IsNullOrEmpty(Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value) = True Then
                                            Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("1", Me.Def_Prec)
                                        End If
                                        IdentifierStr = ObjItem.GetIdentifierUrutan_BuatDGVCellEndEdit(Me.DGVPayment, Me.ColDetailPayment_Identifier.Index, Me.Baru, Me.PrimarySalary, Row.Index)
                                        Me.DGVPayment.Item(Me.ColDetailPayment_Identifier.Index, e.RowIndex).Value = IdentifierStr
                                        Me.DGVPayment.Item(Me.ColDetailPayment_IndexRowSalary.Index, e.RowIndex).Value = Row.Index
                                        BufferPrimaryAkunPay = CDbl(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value)
                                        Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun

                                    End If
                                End If
                            ElseIf ij.mPrimaryAkun = -1 Then
                                If Me.DGVPayment.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_PayAccount = e.RowIndex
                                    Me.BackgroundWorker_Payment.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DGVPayment)
                        '----------------------------------------------------
                    End If
                    If e.ColumnIndex = Me.ColDetailPayment_Amount.Index Then
                        If Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value = "-1" Or _
                        String.IsNullOrEmpty(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value) = True Then
                            MsgBox("Please input account code.", MsgBoxStyle.Information, "Information")
                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Me.EraseRowIndex_PayAccount = e.RowIndex
                            Me.BackgroundWorker_Payment.RunWorkerAsync()
                            Exit Sub
                        End If
                        If Me.ObjCD.CekBilanganPositif(Me.DGVPayment.Item(Me.ColDetailPayment_Amount.Index, e.RowIndex).Value) = False Then
                            MsgBox("Please input positive number.", MsgBoxStyle.Information, "Information")
                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.Def_Prec)
                            Exit Sub
                        End If
                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value, Me.Def_Prec)
                    End If
                    If e.ColumnIndex = Me.ColDetailPayment_Remarks.Index Then
                        If Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value = "-1" Or _
                        String.IsNullOrEmpty(Me.DGVPayment.Item(Me.ColDetailPayment_PrimaryAkun.Index, e.RowIndex).Value) = True Then
                            MsgBox("Please input account code.", MsgBoxStyle.Information, "Information")
                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Me.EraseRowIndex_PayAccount = e.RowIndex
                            Me.BackgroundWorker_Payment.RunWorkerAsync()
                            Exit Sub
                        End If
                        If Me.ObjCD.CekDataKosong(Row.Cells(Me.ColDetailPayment_Remarks.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColDetailPayment_Remarks.Index Then
                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.DGVPayment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_Payment_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_Payment.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_Payment_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_Payment.RunWorkerCompleted
        Try
            Me.DGVPayment.Rows.RemoveAt(Me.EraseRowIndex_PayAccount)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVPayment_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DGVPayment.UserDeletingRow
        Try
            If Me.NonAktif = False Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class