Option Explicit On
Option Strict Off
Imports System.Windows.Forms

Public Class FrmTransaksiCuti_Pengajuan_Setting
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color
    Dim EraseRowIndex = -1
    Dim Buffer As String = ""
    'design notes:
    '1. kelasnya menggunakan ClsTrans_Cuti_Pengajuan + dse nya juga pake hal yang sama
    '2. buat enum pada ClsTrans_Cuti_Pengajuan --> seperti dibawah, stlh itu yg dibawah dihapuskan
    '3. tabel menggunakan: tbltrans_cuti_pengajuan_setting
    '4. modifikasi batasan hapus user dan hapus karyawan, perhitungkan tabel ini juga
    '6. ketika new pada transaksi pengajuan cuti, harus memperhitungkan hak akses orang ybs. jika all employee, maka tidak ada perubahan
    '   jika hanya memiliki 1 akses saja:
    '   - maka secara langsung hanya bisa mengajukan cuti untuk karyawan yang bersangkutan, ketika btn new, langsung load nama karyawan dan 
    '     primary-nya, button find karyawan - enable false dan txt kode karyawan read only
    '   - ketika edit, button find karyawan - enable false dan txt kode karyawan read only
    '   - ketika find hanya bisa mencari pengajuan cuti untuk karyawan ybs saja

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Enum EnumKolom As Integer
        User
        Karyawan
    End Enum

    Private Function GetIndexItemInput(ByVal vPrimary As Double, ByVal vKolom As EnumKolom) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.DgvItem.Rows
                If vKolom = EnumKolom.User Then
                    If rowDGV.Cells(Me.ColPrimaryUser.Index).Value = vPrimary Then
                        Return rowDGV.Index
                    End If
                ElseIf vKolom = EnumKolom.Karyawan Then
                    If rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value = vPrimary Then
                        Return rowDGV.Index
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            If String.IsNullOrEmpty(Me.DgvItem(Me.ColUserCode.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColUserCode.Index Or e.ColumnIndex = Me.ColPrivilage.Index Or _
                e.ColumnIndex = Me.ColRemarks.Index Then
                    Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                ElseIf e.ColumnIndex = Me.ColEmployeeCode.Index Then
                    CboDgv = Me.DgvItem(Me.ColPrivilage.Index, e.RowIndex)
                    If CboDgv.Items.IndexOf(Me.DgvItem.Item(Me.ColPrivilage.Index, e.RowIndex).Value) = _
                    ClsTrans_Cuti_Pengajuan.EnumPrivilagePengajuanCuti.Concerned_Employee Then
                        Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColUserCode.Index Then
                    Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)
        Dim ObjUser As New Core_Module.ClsUser
        Dim ObjKaryawan As New Core_Module.ClsMasterKaryawan
        Dim index As Integer = -1
        Dim Primary As Double = -1
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            With row
                '===user code
                If e.ColumnIndex = Me.ColUserCode.Index Then
                    If String.IsNullOrEmpty(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If ObjUser.ApakahSudahAda(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Primary = ObjUser.GetPrimary(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemInput(Primary, EnumKolom.User)
                            If index < 0 Then
                                Me.DgvItem.Item(Me.ColPrimaryUser.Index, e.RowIndex).Value = Primary
                                Me.DgvItem.Item(Me.ColUserCode.Index, e.RowIndex).Value = ObjUser.GetKodeUser(Primary)
                                Me.DgvItem.Item(Me.ColUserName.Index, e.RowIndex).Value = ObjUser.GetNamaByPrimary(Primary)
                                CboDgv = Me.DgvItem.Item(Me.ColPrivilage.Index, e.RowIndex)
                                Me.DgvItem.Item(Me.ColPrivilage.Index, e.RowIndex).Value = CboDgv.Items.Item(0)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication in user code, please change your code", MsgBoxStyle.Information, "Information")
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterUser
                            ij.TxtCari.Text = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            Primary = ObjUser.GetPrimary(ij.Kodeuser)
                            index = Me.GetIndexItemInput(Primary, EnumKolom.User)
                            If Primary = -1 Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.EraseRowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            Else
                                If index < 0 Then
                                    Me.DgvItem.Item(Me.ColPrimaryUser.Index, e.RowIndex).Value = Primary
                                    Me.DgvItem.Item(Me.ColUserCode.Index, e.RowIndex).Value = ObjUser.GetKodeUser(Primary)
                                    Me.DgvItem.Item(Me.ColUserName.Index, e.RowIndex).Value = ObjUser.GetNamaByPrimary(Primary)
                                    CboDgv = Me.DgvItem.Item(Me.ColPrivilage.Index, e.RowIndex)
                                    Me.DgvItem.Item(Me.ColPrivilage.Index, e.RowIndex).Value = CboDgv.Items.Item(0)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in User code, please change your code", MsgBoxStyle.Information, "Information")
                                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = ObjUser.GetKodeUser(Primary)
                                    End If
                                End If
                            End If
                            ij = Nothing
                        End If
                    Else
                        If e.RowIndex <> Me.DgvItem.Rows.Count - 1 Then
                            Me.EraseRowIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                End If

                '===Privilege
                If e.ColumnIndex = Me.ColPrivilage.Index Then
                    CboDgv = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex)
                    If CboDgv.Items.IndexOf(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = _
                    ClsTrans_Cuti_Pengajuan.EnumPrivilagePengajuanCuti.All_Employee Then
                        Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = -1
                        Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = ""
                        Me.DgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = ""
                    End If
                End If

                '===employee code
                If e.ColumnIndex = Me.ColEmployeeCode.Index Then
                    If String.IsNullOrEmpty(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If ObjKaryawan.ApakahSudahAda(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Primary = ObjKaryawan.GetPrimaryByCode(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemInput(Primary, EnumKolom.Karyawan)
                            If index < 0 Then
                                Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = Primary
                                Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = ObjKaryawan.GetCodeByPrimary(Primary)
                                Me.DgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = ObjKaryawan.GetFullNameByPrimary(Primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication in Employee code, please change your code", MsgBoxStyle.Information, "Information")
                                    Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = -1
                                    Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = ""
                                    Me.DgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = ""
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmMasterKaryawanCari
                            ij.txtCari.Text = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            index = Me.GetIndexItemInput(ij.PrimaryEmploye, EnumKolom.Karyawan)
                            If ij.PrimaryEmploye = -1 Then
                                Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = -1
                                Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = ""
                                Me.DgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = ""
                            Else
                                If index < 0 Then
                                    Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = ObjKaryawan.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.DgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = ObjKaryawan.GetFullNameByPrimary(ij.PrimaryEmploye)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in Employee code, please change your code", MsgBoxStyle.Information, "Information")
                                        Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = -1
                                        Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = ""
                                        Me.DgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = ""
                                        Exit Sub
                                    Else
                                        Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = ObjUser.GetKodeUser(Primary)
                                    End If
                                End If
                            End If
                            ij = Nothing
                        End If
                    Else
                        Me.DgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = -1
                        Me.DgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = ""
                        Me.DgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = ""
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjKaryawan = Nothing
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim ObjLeave As New ClsTrans_Cuti_Pengajuan
        Try
            hasil = MsgBox("Save data ?", MsgBoxStyle.YesNo, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
                If ObjLeave.SimpanBaruSetting(Me, Me.UserCode) = True Then
                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiCuti_Pengajuan_Setting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ObjPengajuanCuti As New ClsTrans_Cuti_Pengajuan
        Try
            Me.ColPrivilage.HeaderText = "Privilege"
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            'isi dgv
            ObjPengajuanCuti.IsiDGVsetting(Me.DgvItem, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItem.RowsAdded
        'Dim rowDGV As DataGridViewRow
        Try
            'Dim CboDgv As DataGridViewComboBoxCell
            'For Each rowDGV In DgvItem.Rows
            '    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColUserCode.Index).Value) = True Then
            '        'CboDgv = rowDGV.Cells(Me.ColPrivilage.Index)
            '        'rowDGV.Cells(Me.ColPrivilage.Index).Value = CboDgv.Items.Item(0)
            '    End If
            'Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvItem.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class