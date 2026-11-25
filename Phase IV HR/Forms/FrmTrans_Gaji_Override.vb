Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing, Core_Module
Imports System.IO.StreamReader
'Imports Excel = Microsoft.Office.Interop.Excel

Public Class FrmTrans_Gaji_Override
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Tipe As ClsTransaksiSalary_Override.EnumType = ClsTransaksiSalary_Override.EnumType.SalaryOverride
    Public LocationFileSumber As String = ""

    Dim Obj As New ClsTransaksiSalary_Override
    Dim ObjItem As New ClsTransaksiSalary_Override_Item
    Dim ObjFormula As New ClsMasterSalaryFormula
    Dim ObjFormulaItem As New Core_Module.ClsMasterSalaryFormula_item
    Dim ObjOther As New Core_Module.ClsTrans_OtherDeduction
    Dim ObjEmp As New Core_Module.ClsMasterKaryawan
    Dim ObjDept As New Core_Module.ClsMasterDepartement
    Dim ObjCurr As New clsMasterCurrency
    Dim ObjCD As New ClsCekData
    Dim ObjUser As New Core_Module.ClsUser

    Dim Baru As Boolean
    Dim Buffer As String
    Dim BufferTransPaymentDate As String
    Dim EraseRowIndex_Item As Integer
    Dim Def_Prec As Integer = ObjCurr.GetDecimal_DefaultCurrency

#Region "Setting TAB"
    'Dim setTab As Boolean = False
    'Private Sub btnFind_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.LostFocus
    '    Try
    '        Me.setTab = False
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub btnExit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.GotFocus
    '    Try
    '        Me.setTab = True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub SplitContainer1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer1.GotFocus
    '    Try
    '        If Me.setTab = False Then
    '            Me.btnExit.Focus()
    '        Else
    '            Me.btnFind.Focus()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
#End Region

    Private Sub LblAdvance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAdvance.LinkClicked
        Try
            Me.ContextMenuStrip1.Show(Me.LblAdvance, 0, 20)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiTicketing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Obj = New ClsTransaksiSalary_Override
            Me.Baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            Me.dgvItem.AllowUserToAddRows = False
            Me.CboFormulaType.SelectedIndex = 0
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            If Me.Tipe = ClsTransaksiSalary_Override.EnumType.OtherPayment Then
                Me.ColOverrideAmount.HeaderText = "Amount"
                Me.ColOverrideReason.Visible = False
                Me.Text = "Other Payment"
            End If
            Me.ColEmployeeCode.Frozen = True
            Me.ColEmployeeName.Frozen = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, True, True, True, True)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.BackColorReset()

            'set dgv
            Me.dgvItem.ReadOnly = False
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.AllowUserToAddRows = True

            'Label
            Me.LblAdvance.Enabled = True
            'Me.lblExport.Enabled = True
            Me.lblImport.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub NonAktifkanDetail()
        Try

            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, True, False, True, False)
            'Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnExit
            Me.BackColorDefault()

            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.AllowUserToAddRows = False

            'Label
            Me.LblAdvance.Enabled = False
            Me.lblImport.Enabled = False
            'Me.lblExport.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorReset()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Color.White, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Color.White, False)
            'Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Color.White, False)
            'Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Color.White, False)
            'Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Color.White, False)
            'Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Color.White, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackColorDefault()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            'Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            'Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
            'Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.TxtReadOnly.BackColor, True)
            'Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.TxtReadOnly.BackColor, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, True)
            'Me.ObjCD.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, True)
            'Me.ObjCD.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, True)
            'Me.ObjCD.KosongkanControl(Me.SplitContainer4.Panel1.Controls, True, True, True)
            'Me.ObjCD.KosongkanControl(Me.SplitContainer4.Panel2.Controls, True, True, True)
            Me.TxtDocNumber.Clear()
            Me.DtpDate.Value = Now
            Me.TxtFormulaCode.Clear()
            Me.TxtFormulaDesc.Clear()
            Me.TxtPrimary.Text = "-1"
            Me.TxtPrimarySalaryFormula.Text = "-1"
            Me.TxtIdentifierSalaryFormula.Text = "-1-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
            Me.CboFormulaType.SelectedIndex = 0
            Me.IsikanSubTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim RowDGV As DataGridViewRow
        Try
            '1. cek data yang harus diisi
            KoleksiNamaData.Add("Document No")
            KoleksiTextBox.Add(Me.TxtDocNumber)
            KoleksiNamaData.Add("Component Name")
            KoleksiTextBox.Add(Me.TxtComponentName)

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '2. harus ada data item
            If Me.dgvItem.Rows.Count <= 1 Then
                MsgBox("Please input data item.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.dgvItem.Rows.Count > 1 Then
                For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                    RowDGV = Me.dgvItem.Rows(i)
                    If String.IsNullOrEmpty(RowDGV.Cells(Me.ColOverrideAmount.Index).Value) = True Then
                        MsgBox("Please input override amount on line" + (RowDGV.Index + 1).ToString + ".", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        RowDGV.Selected = True
                        Return False
                    ElseIf CDec(RowDGV.Cells(Me.ColOverrideAmount.Index).Value) < 0 Then
                        MsgBox("Please specify a non-negative value on line " + (RowDGV.Index + 1).ToString + ".", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        RowDGV.Selected = True
                        Return False
                    End If
                Next
                'For i As Integer = 0 To Me.dgvItem.Rows.Count - 2
                '    RowDGV = Me.dgvItem.Rows(i)
                '    If String.IsNullOrEmpty(RowDGV.Cells(Me.ColOverrideReason.Index).Value) = True Then
                '        MsgBox("Please input override reason.", MsgBoxStyle.Information, "Information")
                '        Return False
                '    End If
                'Next
            End If

            'cek perubahan nomor transaksi
            If Me.Baru = False Then
                If Me.Obj.TanggalTransaksi.Month = Me.DtpDate.Value.Month And _
                Me.Obj.TanggalTransaksi.Year = Me.DtpDate.Value.Year Then
                    Me.TxtDocNumber.Text = Me.Obj.DocumentNo
                Else
                    Dim NoTransaksiLama As String = Obj.GetDocumentNo(Me.TxtPrimary.Text)
                    If NoTransaksiLama <> Me.TxtDocNumber.Text Then
                        If Obj.ApakahSudahAda(Me.TxtDocNumber.Text) = False Then
                            If MsgBox("You just changed transaction date, this will change your transaction " _
                                      + "number and there's a possibility your transaction no cannot be change back." _
                                      + "Continue and accept this change?", _
                                      MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                Me.DtpDate.Value = Me.Obj.TanggalTransaksi
                                Me.TxtDocNumber.Text = Me.Obj.DocumentNo
                                Return False
                            End If
                        End If
                    End If
                End If
            End If

            'cek data yang salah, double dalam kode
            If Me.Baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtDocNumber.Text) = True Then
                    Me.TxtDocNumber.Text = Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD, Me.Tipe)
                End If
            Else
                Dim TipeCodeLama As String = Me.Obj.GetDocumentNo(Me.TxtPrimary.Text)
                If TipeCodeLama <> Me.TxtDocNumber.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtDocNumber.Text) = True Then
                        Me.TxtDocNumber.Text = Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD, Me.Tipe)
                    End If
                End If
                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim OldTimeStamp As Date = Me.Obj.GetTimestamp(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.Obj.Primary, ObjCD, Me.dgvItem, Me.Def_Prec, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            '3. Remarks
            If Me.ObjCD.CekDataKosong(Me.TxtRemarks.Text, False) = False Then
                MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                Me.TxtRemarks.Focus()
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Public Sub SalinKeObjek()
        Try
            With Me.Obj
                .DocumentNo = Me.TxtDocNumber.Text
                .TanggalTransaksi = Me.DtpDate.Value
                If Me.CboFormulaType.SelectedIndex = 0 Then
                    .FormulaType = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula
                Else
                    .FormulaType = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula_Other_Deduction
                End If
                .PrimaryFormula = Me.TxtPrimarySalaryFormula.Text
                .IdentifierFormula = Me.TxtIdentifierSalaryFormula.Text
                .Remarks = Me.TxtRemarks.Text
                .Tipe = Me.Tipe
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtDocNumber.Text = .DocumentNo
                Me.DtpDate.Value = .TanggalTransaksi
                If .FormulaType = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula Then
                    Me.CboFormulaType.SelectedIndex = 0
                Else
                    Me.CboFormulaType.SelectedIndex = 1
                End If
                Me.TxtPrimarySalaryFormula.Text = .PrimaryFormula
                Me.TxtIdentifierSalaryFormula.Text = .IdentifierFormula
                If Me.CboFormulaType.SelectedIndex = 0 Then
                    Me.TxtComponentName.Text = Me.ObjFormulaItem.GetDescComponent(Me.TxtIdentifierSalaryFormula.Text)
                    Me.TxtFormulaCode.Text = Me.ObjFormula.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                    Me.TxtFormulaDesc.Text = Me.ObjFormula.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                Else
                    Me.TxtComponentName.Text = Me.ObjOther.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                    Me.TxtFormulaCode.Text = Me.ObjOther.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                    Me.TxtFormulaDesc.Text = Me.ObjOther.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                End If
                Me.TxtRemarks.Text = .Remarks

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                Me.IsikanSubTotal()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanSubTotal()
        Dim rowDGV As DataGridViewRow
        Dim SubTotal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItem.Rows
                SubTotal += CDec(rowDGV.Cells(Me.ColOverrideAmount.Index).Value)
            Next

            Me.TxtSubTotal.Text = Me.ObjCD.DigitGrouping(SubTotal, Me.ObjCurr.GetDecimal_DefaultCurrency())
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Kontrol Button
    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim PrimaryMin As Double = Double.MaxValue
        Dim PrimaryMax As Double = Double.MinValue
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        'Simpan Perubahan
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Not Me.TxtDocNumber.Text = "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                        Me.TxtDocNumber.Clear()
                    Else
                        Me.Obj.CariDetail(Me.TxtPrimary.Text, Me.ObjCD, _
                        Me.dgvItem, Me.Def_Prec, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.Baru = True
            Obj.ResetObject(Me.dgvItem)
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtDocNumber.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD, Me.Tipe)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.DtpDate.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ijk As New FrmTrans_Gaji_Override_Cari2
        Try
            ijk.BackColorAll = Me.BackColorAll
            ijk.cboCriteria.Items.Clear()
            ijk.txtCari.Text = ""
            ijk.dgvHasil.DataSource = ""
            ijk.ShowIcon = False
            ijk.vTipe = Me.Tipe
            ijk.UserCode = Me.UserCode

            ijk.ShowDialog()
            If Not ijk.mPrimary = -1 Then
                If Me.Obj.CariDetail(ijk.mPrimary, Me.ObjCD, _
                Me.dgvItem, Me.Def_Prec, Me) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
            'Label
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.LblAdvance.Enabled = True
                Else
                    Me.LblAdvance.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtDocNumber.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtDocNumber.Text <> "" Then
                If Me.Obj.ApakahBolehEdit(Me.TxtPrimary.Text, Me.UserCode, True, True) = True Then
                    Me.Baru = False
                    Me.AktifkanDetail()
                    Me.DtpDate.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtDocNumber.Text = "" Then
                Me.btnFind.PerformClick()
                Me.LblAdvance.Enabled = False
            End If

            If Me.TxtDocNumber.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.TxtPrimary.Text, Me.UserCode, True, True, Me) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.dgvItem, Me, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            If Me.dgvItem.AllowUserToAddRows = True Then
                                Me.dgvItem.AllowUserToAddRows = False
                            End If
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtRemarks.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, _
                            Me.dgvItem, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, _
                            Me.dgvItem, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.Enter
        Try
            Me.BufferTransPaymentDate = Me.DtpDate.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpDate.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.BufferTransPaymentDate <> Me.DtpDate.Value.Date Then
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            End If
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
            If Me.TxtRemarks.Enabled = True Then
                If Me.Baru = True Then
                    Me.TxtDocNumber.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD, Me.Tipe)
                Else
                    If Me.DtpDate.Value.Month = Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Month And _
                    Me.DtpDate.Value.Year = Obj.GetTanggalTransaksi(Me.TxtPrimary.Text).Year Then
                        Me.TxtDocNumber.Text = Me.Obj.GetDocumentNo(Me.TxtPrimary.Text)
                    Else
                        If Me.DtpDate.Value.Month <> Me.Obj.TanggalTransaksi.Month Or _
                        Me.DtpDate.Value.Year <> Me.Obj.TanggalTransaksi.Year Then
                            Me.TxtDocNumber.Text = Me.Obj.GetNoBaru(Me.DtpDate.Value, Me.ObjCD, Me.Tipe)
                        Else
                            Me.TxtDocNumber.Text = Me.Obj.GetDocumentNo(Me.TxtPrimary.Text)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemInput(ByVal vPrimaryEmployee As Double, ByVal RowIndex As Integer) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For ij As Integer = 0 To Me.dgvItem.Rows.Count - 1
                rowDGV = Me.dgvItem.Rows(ij)
                If ij <> RowIndex Then
                    If rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value = vPrimaryEmployee Then
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

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColEmployeeCode.Index Or _
            e.ColumnIndex = Me.ColOverrideAmount.Index Then
                If e.ColumnIndex = Me.ColOverrideAmount.Index Then
                    If Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = "-1" _
                    Or String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value) = True Then
                        e.Cancel = True
                    Else
                        Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            Else
                If e.ColumnIndex = Me.ColOverrideReason.Index Then
                    If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColOverrideAmount.Index, e.RowIndex).Value) = True Then
                        e.Cancel = True
                    Else
                        If Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = "-1" _
                         Or String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value) = True Then
                            e.Cancel = True
                        Else
                            Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    End If
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Dim Index As Integer
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColOverrideReason.Index Then
                        If Me.ObjCD.CekDataKosong(Row.Cells(Me.ColOverrideReason.Index).Value, False) = False Then
                            MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    If e.ColumnIndex = Me.ColEmployeeCode.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjEmp.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            Index = Me.GetIndexItemInput(Me.ObjEmp.GetPrimaryByCode(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), e.RowIndex)
                            If Index < 0 Then
                                Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = Me.ObjEmp.GetPrimaryByCode(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColJobPositionCode.Index, e.RowIndex).Value = Me.ObjEmp.GetPositionCode(Me.dgvItem(Me.ColPrimaryEmployee.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColJobPositionDescription.Index, e.RowIndex).Value = Me.ObjEmp.GetPositionDescription(Me.dgvItem(Me.ColPrimaryEmployee.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColDepartmentCode.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(Me.ObjEmp.GetDeptByPrimary(Me.dgvItem(Me.ColPrimaryEmployee.Index, e.RowIndex).Value))
                                Me.dgvItem.Item(Me.ColDepartmentDesc.Index, e.RowIndex).Value = Me.ObjEmp.GetDeptDescription(Me.dgvItem(Me.ColPrimaryEmployee.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColOverrideAmount.Index, e.RowIndex).Value = "0"
                                Me.dgvItem.Item(Me.ColOverrideReason.Index, e.RowIndex).Value = ""
                            Else
                                If e.RowIndex <> Index Then
                                    MsgBox("Duplication employee code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_Item = e.RowIndex
                                        Me.BackgroundWorker_DetailEmployee.RunWorkerAsync()
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New Core_Module.FrmMasterKaryawanCari
                            ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryEmploye <> -1 Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItemInput(ij.PrimaryEmploye, e.RowIndex)
                                If Index < 0 Then
                                    Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    Me.dgvItem.Item(Me.ColEmployeeCode.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColEmployeeName.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColJobPositionCode.Index, e.RowIndex).Value = Me.ObjEmp.GetPositionCode(Me.dgvItem(Me.ColPrimaryEmployee.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColJobPositionDescription.Index, e.RowIndex).Value = Me.ObjEmp.GetPositionDescription(Me.dgvItem(Me.ColPrimaryEmployee.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColDepartmentCode.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(Me.ObjEmp.GetDeptByPrimary(Me.dgvItem(Me.ColPrimaryEmployee.Index, e.RowIndex).Value))
                                    Me.dgvItem.Item(Me.ColDepartmentDesc.Index, e.RowIndex).Value = Me.ObjEmp.GetDeptDescription(Me.dgvItem(Me.ColPrimaryEmployee.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColOverrideAmount.Index, e.RowIndex).Value = "0"
                                    Me.dgvItem.Item(Me.ColOverrideReason.Index, e.RowIndex).Value = ""
                                Else
                                    If e.RowIndex <> Index Then
                                        MsgBox("Duplication employee code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_Item = e.RowIndex
                                            Me.BackgroundWorker_DetailEmployee.RunWorkerAsync()
                                        Else
                                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvItem)
                                        Exit Sub
                                    End If
                                End If
                            ElseIf ij.PrimaryEmploye = -1 Then
                                If Me.dgvItem.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_Item = e.RowIndex
                                    Me.BackgroundWorker_DetailEmployee.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvItem)
                    End If
                    'Batasan tidak boleh negatif untuk override amount
                    If e.ColumnIndex = Me.ColOverrideAmount.Index Then
                        If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please input positive number.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.Def_Prec)
                    End If
                Else
                    If e.ColumnIndex = Me.ColOverrideReason.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            End With

            Me.IsikanSubTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_DetailEmployee_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_DetailEmployee.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_DetailEmployee_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_DetailEmployee.RunWorkerCompleted
        Try
            Me.dgvItem.Rows.RemoveAt(Me.EraseRowIndex_Item)
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BersihkanDGVEmployee(ByVal vBufferPrimaryEmp As Double, ByVal vPrimaryEmp As Double, ByVal vBufferIdentifier As String, ByVal vIdentifier As String)
        Try
            If vPrimaryEmp = -1 Then
                Me.dgvItem.Rows.Clear()
            End If
            If vBufferPrimaryEmp <> vPrimaryEmp Or _
            vBufferIdentifier <> vIdentifier Then
                Me.dgvItem.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindComponentName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindComponentName.Click
        Dim i As New FrmTrans_Gaji_Override_Cari_Formula
        Dim ObjCode As New Core_Module.ClsMasterSalaryFormula
        Dim ObjCode_1 As New Core_Module.ClsTrans_OtherDeduction
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtComponentName.Text <> "" Then
                i.txtCari.Text = Me.TxtComponentName.Text
            End If
            If Me.CboFormulaType.SelectedIndex = 0 Then
                i.TypePencarian = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula
                i.mPrimaryUser = -1
                i.ShowDialog()
                If i.mPrimary > 0 Then
                    'Me.BersihkanDGVEmployee(Me.TxtPrimarySalaryFormula.Text, i.mPrimary, _
                    'Me.TxtIdentifierSalaryFormula.Text, i.mIdentifierFormula)
                    Me.TxtPrimarySalaryFormula.Text = i.mPrimary
                    Me.TxtComponentName.Text = Me.ObjFormulaItem.GetDescComponent(i.mIdentifierFormula)
                    Me.TxtIdentifierSalaryFormula.Text = i.mIdentifierFormula
                    Me.TxtFormulaCode.Text = ObjCode.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                    Me.TxtFormulaDesc.Text = ObjCode.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                    Me.TxtRemarks.Focus()
                Else
                    Me.TxtFormulaCode.Clear()
                    Me.TxtFormulaDesc.Clear()
                    Me.TxtPrimarySalaryFormula.Text = "-1"
                    Me.TxtIdentifierSalaryFormula.Text = "-1-1"
                    Me.TxtComponentName.Clear()
                    Me.TxtComponentName.Focus()
                    Me.dgvItem.Rows.Clear()
                End If
            Else
                i.TypePencarian = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula_Other_Deduction
                i.ShowDialog()
                If i.mPrimary > 0 Then
                    'Me.BersihkanDGVEmployee(Me.TxtPrimarySalaryFormula.Text, i.mPrimary, _
                    'Me.TxtIdentifierSalaryFormula.Text, i.mPrimary)
                    Me.TxtPrimarySalaryFormula.Text = i.mPrimary
                    Me.TxtComponentName.Text = ObjCode_1.GetCodeByPrimary(i.mPrimary)
                    Me.TxtIdentifierSalaryFormula.Text = i.mPrimary
                    Me.TxtFormulaCode.Text = ObjCode_1.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                    Me.TxtFormulaDesc.Text = ObjCode_1.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                    Me.TxtRemarks.Focus()
                Else
                    Me.TxtFormulaCode.Clear()
                    Me.TxtFormulaDesc.Clear()
                    Me.TxtPrimarySalaryFormula.Text = "-1"
                    Me.TxtIdentifierSalaryFormula.Text = "-1-1"
                    Me.TxtComponentName.Clear()
                    Me.TxtComponentName.Focus()
                    Me.dgvItem.Rows.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            ObjCode = Nothing
            ObjCode_1 = Nothing
        End Try
    End Sub

    Private Sub TxtComponentName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtComponentName.KeyDown
        Dim i As New FrmTrans_Gaji_Override_Cari_Formula
        Dim ObjCode As New Core_Module.ClsMasterSalaryFormula
        Dim ObjCode_1 As New Core_Module.ClsTrans_OtherDeduction
        i.BackColorAll = Me.BackColorAll
        Try
            If e.KeyValue = 13 Then
                If Me.TxtComponentName.Text <> "" Then
                    If Me.CboFormulaType.SelectedIndex = 0 Then
                        If Me.Obj.GetJmlByComponentName(Me.TxtComponentName.Text, Me.ObjUser.GetPrimary(Me.UserCode)) > 1 Then
                            Me.BtnFindComponentName.PerformClick()
                        Else
                            If Me.Obj.ApakahSudahAda_FormulaSalary(Me.TxtComponentName.Text, Me.ObjUser.GetPrimary(Me.UserCode)) = True Then
                                'Me.BersihkanDGVEmployee(Me.TxtPrimarySalaryFormula.Text, Me.ObjFormulaItem.GetPrimaryFormula(Me.TxtComponentName.Text), _
                                'Me.TxtIdentifierSalaryFormula.Text, Me.ObjFormulaItem.GetIdentifierFormula(Me.TxtComponentName.Text))
                                Me.TxtPrimarySalaryFormula.Text = Me.ObjFormulaItem.GetPrimaryFormula(Me.TxtComponentName.Text)
                                Me.TxtIdentifierSalaryFormula.Text = Me.ObjFormulaItem.GetIdentifierFormula(Me.TxtComponentName.Text)
                                Me.TxtFormulaCode.Text = ObjCode.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                                Me.TxtFormulaDesc.Text = ObjCode.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                                Me.TxtRemarks.Focus()
                            Else
                                Me.BtnFindComponentName.PerformClick()
                            End If
                        End If
                    Else
                        If ObjCode_1.ApakahSudahAda(Me.TxtComponentName.Text) = True Then
                            'Me.BersihkanDGVEmployee(Me.TxtPrimarySalaryFormula.Text, ObjCode_1.GetPrimary(Me.TxtComponentName.Text), _
                            'Me.TxtIdentifierSalaryFormula.Text, Me.TxtPrimarySalaryFormula.Text)
                            Me.TxtPrimarySalaryFormula.Text = ObjCode_1.GetPrimary(Me.TxtComponentName.Text)
                            Me.TxtIdentifierSalaryFormula.Text = Me.TxtPrimarySalaryFormula.Text
                            Me.TxtFormulaCode.Text = ObjCode_1.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                            Me.TxtFormulaDesc.Text = ObjCode_1.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                            Me.TxtRemarks.Focus()
                        Else
                            Me.BtnFindComponentName.PerformClick()
                        End If
                    End If
                Else
                    Me.BtnFindComponentName.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCode = Nothing
            ObjCode_1 = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtComponentName_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtComponentName.Leave
        Dim i As New FrmTrans_Gaji_Override_Cari_Formula
        Dim ObjCode As New Core_Module.ClsMasterSalaryFormula
        Dim ObjCode_1 As New Core_Module.ClsTrans_OtherDeduction
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtComponentName.Text <> "" Then
                If Me.CboFormulaType.SelectedIndex = 0 Then
                    If Me.Obj.GetJmlByComponentName(Me.TxtComponentName.Text, Me.ObjUser.GetPrimary(Me.UserCode)) > 1 Then
                        If Me.Obj.ApakahSudahAda_IdentifierFormula(Me.TxtIdentifierSalaryFormula.Text, Me.ObjUser.GetPrimary(Me.UserCode)) = True Then
                            'Me.BersihkanDGVEmployee(Me.TxtPrimarySalaryFormula.Text, Me.ObjFormulaItem.GetPrimaryFormula_Identifier(Me.TxtIdentifierSalaryFormula.Text), _
                            'Me.TxtIdentifierSalaryFormula.Text, Me.TxtIdentifierSalaryFormula.Text)
                            Me.TxtPrimarySalaryFormula.Text = Me.ObjFormulaItem.GetPrimaryFormula_Identifier(Me.TxtIdentifierSalaryFormula.Text)
                            Me.TxtFormulaCode.Text = ObjCode.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                            Me.TxtFormulaDesc.Text = ObjCode.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                        Else
                            If Me.TxtComponentName.Text <> "" Then
                                i.txtCari.Text = Me.TxtComponentName.Text
                            End If
                            i.TypePencarian = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula
                            i.mPrimaryUser = -1
                            i.ShowDialog()
                            If i.mPrimary > 0 Then
                                'Me.BersihkanDGVEmployee(Me.TxtPrimarySalaryFormula.Text, i.mPrimary, _
                                'Me.TxtIdentifierSalaryFormula.Text, i.mIdentifierFormula)
                                Me.TxtPrimarySalaryFormula.Text = i.mPrimary
                                Me.TxtComponentName.Text = Me.ObjFormulaItem.GetDescComponent(i.mIdentifierFormula)
                                Me.TxtIdentifierSalaryFormula.Text = Me.ObjFormulaItem.GetIdentifierFormula(Me.TxtComponentName.Text)
                                Me.TxtFormulaCode.Text = ObjCode.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                                Me.TxtFormulaDesc.Text = ObjCode.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                            Else
                                Me.TxtFormulaCode.Clear()
                                Me.TxtFormulaDesc.Clear()
                                Me.TxtPrimarySalaryFormula.Text = "-1"
                                Me.TxtIdentifierSalaryFormula.Text = "-1-1"
                                Me.TxtComponentName.Clear()
                                Me.dgvItem.Rows.Clear()
                            End If
                        End If
                    Else
                        'Kondisi Normal
                        If Me.Obj.ApakahSudahAda_FormulaSalary(Me.TxtComponentName.Text, Me.ObjUser.GetPrimary(Me.UserCode)) = True Then
                            'Me.BersihkanDGVEmployee(Me.TxtPrimarySalaryFormula.Text, Me.ObjFormulaItem.GetPrimaryFormula(Me.TxtComponentName.Text), _
                            'Me.TxtIdentifierSalaryFormula.Text, Me.ObjFormulaItem.GetIdentifierFormula(Me.TxtComponentName.Text))
                            Me.TxtPrimarySalaryFormula.Text = Me.ObjFormulaItem.GetPrimaryFormula(Me.TxtComponentName.Text)
                            Me.TxtIdentifierSalaryFormula.Text = Me.ObjFormulaItem.GetIdentifierFormula(Me.TxtComponentName.Text)
                            Me.TxtFormulaCode.Text = ObjCode.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                            Me.TxtFormulaDesc.Text = ObjCode.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                        Else
                            If Me.TxtComponentName.Text <> "" Then
                                i.txtCari.Text = Me.TxtComponentName.Text
                            End If
                            i.TypePencarian = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula
                            i.ShowDialog()
                            If i.mPrimary > 0 Then
                                'Me.BersihkanDGVEmployee(Me.TxtPrimarySalaryFormula.Text, i.mPrimary, _
                                'Me.TxtIdentifierSalaryFormula.Text, i.mIdentifierFormula)
                                Me.TxtPrimarySalaryFormula.Text = i.mPrimary
                                Me.TxtComponentName.Text = Me.ObjFormulaItem.GetDescComponent(i.mIdentifierFormula)
                                Me.TxtIdentifierSalaryFormula.Text = Me.ObjFormulaItem.GetIdentifierFormula(Me.TxtComponentName.Text)
                                Me.TxtFormulaCode.Text = ObjCode.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                                Me.TxtFormulaDesc.Text = ObjCode.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                            Else
                                Me.TxtFormulaCode.Clear()
                                Me.TxtFormulaDesc.Clear()
                                Me.TxtPrimarySalaryFormula.Text = "-1"
                                Me.TxtIdentifierSalaryFormula.Text = "-1-1"
                                Me.TxtComponentName.Clear()
                                Me.dgvItem.Rows.Clear()
                            End If
                        End If
                    End If
                Else
                    If ObjCode_1.ApakahSudahAda(Me.TxtComponentName.Text) = True Then
                        'Me.BersihkanDGVEmployee(Me.TxtPrimarySalaryFormula.Text, ObjCode_1.GetPrimary(Me.TxtComponentName.Text), _
                        'Me.TxtIdentifierSalaryFormula.Text, Me.TxtPrimarySalaryFormula.Text)
                        Me.TxtPrimarySalaryFormula.Text = ObjCode_1.GetPrimary(Me.TxtComponentName.Text)
                        Me.TxtIdentifierSalaryFormula.Text = ObjCode_1.GetPrimary(Me.TxtComponentName.Text)
                        Me.TxtFormulaCode.Text = ObjCode_1.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                        Me.TxtFormulaDesc.Text = ObjCode_1.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                    Else
                        If Me.TxtComponentName.Text <> "" Then
                            i.txtCari.Text = Me.TxtComponentName.Text
                        End If
                        i.TypePencarian = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula_Other_Deduction
                        i.ShowDialog()
                        If i.mPrimary > 0 Then
                            'Me.BersihkanDGVEmployee(Me.TxtPrimarySalaryFormula.Text, i.mPrimary, _
                            'Me.TxtIdentifierSalaryFormula.Text, i.mPrimary)
                            Me.TxtPrimarySalaryFormula.Text = i.mPrimary
                            Me.TxtComponentName.Text = ObjCode_1.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                            Me.TxtIdentifierSalaryFormula.Text = Me.TxtPrimarySalaryFormula.Text
                            Me.TxtFormulaCode.Text = ObjCode_1.GetCodeByPrimary(Me.TxtPrimarySalaryFormula.Text)
                            Me.TxtFormulaDesc.Text = ObjCode_1.GetDescription(Me.TxtPrimarySalaryFormula.Text)
                        Else
                            Me.TxtFormulaCode.Clear()
                            Me.TxtFormulaDesc.Clear()
                            Me.TxtPrimarySalaryFormula.Text = "-1"
                            Me.TxtIdentifierSalaryFormula.Text = "-1-1"
                            Me.TxtComponentName.Clear()
                            Me.dgvItem.Rows.Clear()
                        End If
                    End If
                End If
            Else
                Me.TxtFormulaCode.Clear()
                Me.TxtFormulaDesc.Clear()
                Me.TxtPrimarySalaryFormula.Text = "-1"
                Me.TxtIdentifierSalaryFormula.Text = "-1-1"
                Me.TxtComponentName.Clear()
                Me.dgvItem.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            ObjCode = Nothing
            ObjCode_1 = Nothing
        End Try
    End Sub

    Private Sub TxtComponentName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtComponentName.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtComponentName.Enabled = True Then
                    If Me.TxtComponentName.ReadOnly = False Then
                        Me.TxtComponentName.BackColor = Me.BackColorAll
                    Else
                        Me.TxtComponentName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtComponentName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtComponentName.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtComponentName.Enabled = True Then
                    If Me.TxtComponentName.ReadOnly = False Then
                        Me.TxtComponentName.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtComponentName.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.Enabled = True Then
                    If Me.TxtRemarks.ReadOnly = False Then
                        Me.TxtRemarks.BackColor = Me.BackColorAll
                    Else
                        Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtRemarks.Enabled = True Then
                    If Me.TxtRemarks.ReadOnly = False Then
                        Me.TxtRemarks.BackColor = Me.TxtPrimary.BackColor
                    Else
                        Me.TxtRemarks.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyByDepartementToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ApplyByDepartementToolStripMenuItem.Click
        Dim Frm As New FrmMasterDepartementCari
        Try
            Frm.BackColorAll = Me.BackColorAll
            Frm.ShowDialog()
            If Frm.PrimaryDepartement <> -1 Then
                Me.Obj.CariDataKaryawan(ClsMasterCuti.enumCariData.dept, Frm.PrimaryDepartement, Me.Def_Prec, Me)
            End If
            Me.IsikanSubTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyToAJobPositionToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ApplyToAJobPositionToolStripMenuItem.Click
        Dim Frm As New FrmMasterJobPosition_Cari
        Try
            Frm.BackColorAll = Me.BackColorAll
            Frm.ShowDialog()
            If Frm.mPrimary <> -1 Then
                Me.Obj.CariDataKaryawan(ClsMasterCuti.enumCariData.Job, Frm.mPrimary, Me.Def_Prec, Me)
            End If
            Me.IsikanSubTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ApplyToAllEmployeeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ApplyToAllEmployeeToolStripMenuItem.Click
        Try
            Me.Obj.CariDataKaryawan(ClsMasterCuti.enumCariData.AllEmp, -1, Me.Def_Prec, Me)
            Me.IsikanSubTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CopyFromOtherDataToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CopyFromOtherDataToolStripMenuItem.Click
        Dim Frm As New FrmTrans_Gaji_Override_Cari2
        Try
            Frm.BackColorAll = Me.BackColorAll
            Frm.vTipe = Me.Tipe
            Frm.UserCode = Me.UserCode

            Frm.ShowDialog()
            If Frm.mPrimary <> -1 Then
                Me.Obj.CariDataKaryawan(ClsMasterCuti.enumCariData.Other, Frm.mPrimary, Me.Def_Prec, Me)
            End If
            Me.IsikanSubTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Dim ij As New FrmTrans_Gaji_Override_Cetak
        Try
            If Me.TxtDocNumber.Text = "" Then
                Me.btnFind.PerformClick()
                Me.LblAdvance.Enabled = False
            End If

            If Me.TxtDocNumber.Text <> "" Then
                ij.mPrimary = Me.TxtPrimary.Text
                ij.mDef_Prec = Me.Def_Prec
                ij.ShowIcon = False
                ij.vTipe = Me.Tipe
                ij.ShowDialog()
                ij.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub LblAdvance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblAdvance.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.dgvItem.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboFormulaType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboFormulaType.SelectedIndexChanged
        Try
            Me.TxtFormulaCode.Clear()
            Me.TxtFormulaDesc.Clear()
            Me.TxtPrimarySalaryFormula.Text = "-1"
            Me.TxtIdentifierSalaryFormula.Text = "-1-1"
            Me.TxtComponentName.Clear()
            Me.TxtComponentName.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.IsikanSubTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Openfile(ByRef vFileExtStr As String) As Boolean
        Dim FileExtStr As String = ""

        Try
            '1. cari file csv format infinity
            Me.OpenFileDialog1.Title = "Please specify file"
            Me.OpenFileDialog1.Filter = "csv files (*.csv)|*.csv"
            'Me.OpenFileDialog1.Filter = "Excel Worksheets (*.xlsx)|*.xlsx|excel files (*.xls)|*.xls|csv files (*.csv)|*.csv"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "csv"
            Me.OpenFileDialog1.CheckFileExists = True
            Me.OpenFileDialog1.Multiselect = False
            Me.OpenFileDialog1.FileName = ""
            If Me.OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.OK Then
                Return False
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Return False
            End If

            Me.LocationFileSumber = Me.OpenFileDialog1.FileName
            vFileExtStr = System.IO.Path.GetExtension(Me.OpenFileDialog1.FileName)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetFrmDia(ByRef vFileExtStr As String) As Boolean
        Try
            Me.FrmDia.Title = "Please specify file"
            Me.FrmDia.Filter = "csv files (*.csv)|*.csv"
            'Me.FrmDia.Filter = "Excel Worksheets (*.xlsx)|*.xlsx|excel files (*.xls)|*.xls|csv files (*.csv)|*.csv"
            Me.FrmDia.FilterIndex = 1
            'Me.FrmDia.DefaultExt = "xls"
            Me.FrmDia.DefaultExt = "csv"
            If Me.Tipe = ClsTransaksiSalary_Override.EnumType.SalaryOverride Then
                Me.FrmDia.FileName = "Override " & Me.TxtComponentName.Text
                Me.FrmDia.Title = "Save override data"
            Else
                Me.FrmDia.FileName = "Other Payment " & Me.TxtComponentName.Text
                Me.FrmDia.Title = "Save other payment data"
            End If

            If Me.FrmDia.ShowDialog <> Windows.Forms.DialogResult.OK And Me.FrmDia.FileName.Trim.Length = 0 Then
                Return False
            End If

            vFileExtStr = System.IO.Path.GetExtension(Me.FrmDia.FileName)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Enum EnumFormatFile As Integer
        CsvFormat
        ExcelFormat
    End Enum

    Private Sub lblImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblImport.Click
        Dim FileExtStr As String = ""

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If Me.Openfile(FileExtStr) = False Then
                Exit Sub
            End If

            If FileExtStr.StartsWith(".csv") Then
                Me.ImportFile(EnumFormatFile.CsvFormat)
                'ElseIf FileExtStr.StartsWith(".xlsx") Then
                '    Me.ImportFile(EnumFormatFile.ExcelFormat)
                'ElseIf FileExtStr.StartsWith(".xls") Then
                '    Me.ImportFile(EnumFormatFile.ExcelFormat)
            End If

            Me.IsikanSubTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub lblExport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblExport.Click
        Dim DTAEmp As New DseTransaksiSalary_OverrideTableAdapters.tbl_EmployeeTableAdapter
        Dim TBLEmp As New DseTransaksiSalary_Override.tbl_EmployeeDataTable
        Dim RowEmp As DseTransaksiSalary_Override.tbl_EmployeeRow
        Dim listEmployee As New List(Of String)
        Dim ix As Integer = 0
        Dim FileExtStr As String = ""
        Dim hj As New FrmMassageCSV
        Try
            If Me.dgvItem.RowCount < 1 Then
                Exit Sub
            End If

            If Me.dgvItem.Rows.Count = 1 Then
                Dim rD As DataGridViewRow = Me.dgvItem.Rows(0)
                If String.IsNullOrEmpty(rD.Cells(Me.ColEmployeeCode.Index).Value) = True Then
                    MsgBox("There is no data will be exported.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.Cursor = Cursors.WaitCursor
            If Me.GetFrmDia(FileExtStr) = False Then
                Exit Sub
            End If

            DTAEmp.Connection.ConnectionString = Core_Module.GetConstring
            DTAEmp.Fill(TBLEmp)

            For Each rowd As DataGridViewRow In Me.dgvItem.Rows
                If rowd.Cells(Me.ColEmployeeName.Index).Value <> "" Then
                    TBLEmp.DefaultView.RowFilter = "[Primary] = " & rowd.Cells(Me.ColPrimaryEmployee.Index).Value
                    If TBLEmp.DefaultView.Count > 0 Then
                        RowEmp = TBLEmp.DefaultView(0).Row
                        RowEmp.Amount = CDec(rowd.Cells(Me.ColOverrideAmount.Index).Value)
                        RowEmp.Urutan = CInt(rowd.Cells(Me.ColNo.Index).Value)
                    End If
                End If
            Next

            While ix <= TBLEmp.Rows.Count - 1
                RowEmp = TBLEmp.Rows(ix)
                If RowEmp.Amount = 0 Then
                    TBLEmp.Rows.Remove(RowEmp)
                Else
                    ix += 1
                End If
            End While

            If FileExtStr.StartsWith(".csv") Then
                If hj.ShowDialog = Windows.Forms.DialogResult.Yes Then
                    If hj.Rbo = 0 Then
                        Me.ExportFileCsv(TBLEmp)
                    Else
                        Me.ExportFileCsv2(TBLEmp)
                    End If
                End If

                'ElseIf FileExtStr.StartsWith(".xlsx") Then
                '    Me.ExportFileExcel(TBLEmp)
                'ElseIf FileExtStr.StartsWith(".xls") Then
                '    Me.ExportFileExcel(TBLEmp)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    'Taufik
    Private Sub ImportFile(ByVal vEnumFormat As EnumFormatFile)
        Dim SR As System.IO.StreamReader = New System.IO.StreamReader(Me.LocationFileSumber)
        Dim line As String = SR.ReadLine()
        'Dim strArray As String() = line.Split(";"c)
        Dim Tbl As DataTable = New DataTable
        Dim row As DataRow

        Dim NamaFilter As String = ""
        Dim JmlKaryawan As Integer = 0
        Dim ListNama_InDGV As New List(Of String)
        Dim ListNama_1 As New List(Of String) 'list karyawan di dgv yang ga ada di csv
        Dim ListNama_2 As New List(Of String) 'list karyawan di csv yang ga ada di dgv
        Dim StrMsgBox_1 As String = ""
        Dim StrMsgBox_2 As String = ""
        Dim StrMsgBox_3 As String = ""
        Dim ApaExitSub As Boolean = False
        Dim iCOunt As String() = line.Split(","c)
        Dim Icount2 As String() = line.Split(";"c)
        Dim strArray As String()
        Try
            If iCOunt.Length > Icount2.Length Then
                strArray = line.Split(","c)
            Else
                strArray = line.Split(";"c)
            End If
            'pindahkan data csv ke dalam data table

            For Each s As String In strArray
                Tbl.Columns.Add(New DataColumn())
            Next
            Tbl.Columns.Add(New DataColumn())

            If Tbl.Columns.Count <> 3 Then
                MsgBox("Make sure csv file consist of name and amount.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Do
                line = SR.ReadLine
                If Not line = String.Empty Then
                    If iCOunt.Length > Icount2.Length Then
                        row = Tbl.NewRow
                        row.ItemArray = line.Split(","c)
                        row.Item(2) = row.Item(0).ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper
                        Tbl.Rows.Add(row)
                    Else
                        row = Tbl.NewRow
                        row.ItemArray = line.Split(";"c)
                        row.Item(2) = row.Item(0).ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper
                        Tbl.Rows.Add(row)
                    End If
                Else
                    Exit Do
                End If
            Loop

            'pindahkan list karyawan yang ada pada dgv ke tabel csv
            For Each rowd As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowd.Cells(Me.ColEmployeeName.Index).Value) = False Then
                    NamaFilter = rowd.Cells(Me.ColEmployeeName.Index).Value.ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper
                    If ListNama_InDGV.Contains(NamaFilter) = False Then
                        ListNama_InDGV.Add(NamaFilter)
                        'pastikan hanya pindahkan apabila karyawan tsb ada di list karyawan pada file csv
                        Tbl.DefaultView.RowFilter = "[Column3] = '" & NamaFilter.ToString & "'"
                        If Tbl.DefaultView.Count > 0 Then
                            row = Tbl.DefaultView(0).Row
                            rowd.Cells(Me.ColOverrideAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Item(1), Me.Def_Prec)
                        Else
                            If ListNama_1.Contains(rowd.Cells(Me.ColEmployeeName.Index).Value) = False Then
                                ListNama_1.Add(rowd.Cells(Me.ColEmployeeName.Index).Value)
                            End If
                        End If
                    End If
                End If
            Next

            For Each row In Tbl.Rows
                If ListNama_InDGV.Contains(row.Item(2).ToString.Replace(" ", "").Replace(".", "").Replace(",", "").ToUpper) = False Then
                    If ListNama_2.Contains(row.Item(0)) = False Then
                        ListNama_2.Add(row.Item(0))
                    End If
                End If
            Next
            If ListNama_1.Count > 0 Then
                ListNama_1.Sort()
                For Each NamaKary As String In ListNama_1
                    If StrMsgBox_1 = "" Then
                        StrMsgBox_1 = "Employee not found in csv file." & vbCrLf & NamaKary
                    Else
                        StrMsgBox_1 &= ", " & NamaKary
                    End If
                Next
            End If
            If ListNama_2.Count > 0 Then
                ListNama_2.Sort()
                For Each NamaKary As String In ListNama_2
                    If StrMsgBox_2 = "" Then
                        StrMsgBox_2 = "Employee not found in item" & vbCrLf & NamaKary
                    Else
                        StrMsgBox_2 &= ", " & NamaKary
                    End If
                Next
                StrMsgBox_2 &= vbCrLf
            End If

            If StrMsgBox_1 <> "" Or StrMsgBox_2 <> "" Then
                If StrMsgBox_1 <> "" Then
                    StrMsgBox_1 &= vbCrLf & vbCrLf
                End If
                MsgBox(StrMsgBox_1 & StrMsgBox_2, MsgBoxStyle.Information, "Information")
            End If

            MsgBox("File Import Success", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Taufik
    Private Sub ExportFileCsv(ByVal TBLEmp As DseTransaksiSalary_Override.tbl_EmployeeDataTable)
        Dim RowEmp As DseTransaksiSalary_Override.tbl_EmployeeRow
        Dim csvVal As String = ""

        Try
            If Me.dgvItem.Rows.Count <= 0 Then
                MsgBox("There is no data will be exported.", MsgBoxStyle.Information, "Information")
            Else
                csvVal = "NIK;Employee Code;Employee Name;Departement Code;Departement Description;Job Position Code;Job Position Description;Amount"

                TBLEmp.DefaultView.RowFilter = Nothing
                TBLEmp.DefaultView.Sort = "[Urutan] ASC"
                For i As Integer = 0 To TBLEmp.DefaultView.Count - 1
                    RowEmp = TBLEmp.DefaultView(i).Row
                    With RowEmp
                        csvVal += vbNewLine & .NIK.ToString.Replace(",", "") & ";"
                        csvVal += .Code.ToString.Replace(",", "") & ";"
                        csvVal += .Full_Name.ToString.Replace(",", "") & ";"
                        csvVal += .Dept_Code.ToString.Replace(",", "") & ";"
                        csvVal += .Dept_Desc.ToString.Replace(",", "") & ";"
                        csvVal += .Job_Code.ToString.Replace(",", "") & ";"
                        csvVal += .Job_Desc.ToString.Replace(",", "") & ";"
                        csvVal += .Amount.ToString.Replace(",", "")
                    End With
                Next

                My.Computer.FileSystem.WriteAllText(Me.FrmDia.FileName, csvVal, False)
                MsgBox("File Export Success", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '
    Private Sub ExportFileCsv2(ByVal TBLEmp As DseTransaksiSalary_Override.tbl_EmployeeDataTable)
        Dim RowEmp As DseTransaksiSalary_Override.tbl_EmployeeRow
        Dim csvVal As String = ""

        Try
            If Me.dgvItem.Rows.Count <= 0 Then
                MsgBox("There is no data will be exported.", MsgBoxStyle.Information, "Information")
            Else
                csvVal = "NIK, Employee Code, Employee Name, Departement Code, Departement Description, Job Position Code, Job Position Description, Amount"

                TBLEmp.DefaultView.RowFilter = Nothing
                TBLEmp.DefaultView.Sort = "[Urutan] ASC"
                For i As Integer = 0 To TBLEmp.DefaultView.Count - 1
                    RowEmp = TBLEmp.DefaultView(i).Row
                    With RowEmp
                        csvVal += vbNewLine & .NIK.ToString.Replace(";", "") & ","
                        csvVal += .Code.ToString.Replace(";", "") & ","
                        csvVal += .Full_Name.ToString.Replace(";", "") & ","
                        csvVal += .Dept_Code.ToString.Replace(";", "") & ","
                        csvVal += .Dept_Desc.ToString.Replace(";", "") & ","
                        csvVal += .Job_Code.ToString.Replace(";", "") & ","
                        csvVal += .Job_Desc.ToString.Replace(";", "") & ","
                        csvVal += .Amount.ToString.Replace(";", "")
                    End With
                Next

                My.Computer.FileSystem.WriteAllText(Me.FrmDia.FileName, csvVal, False)
                MsgBox("File Export Success", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '

    Private Sub ExportFileExcel(ByVal TBLEmp As DseTransaksiSalary_Override.tbl_EmployeeDataTable)
        'Dim RowEmp As DseTransaksiSalary_Override.tbl_EmployeeRow
        'Dim _excel As New Excel.Application
        'Dim wBook As Excel.Workbook
        'Dim wSheet As Excel.Worksheet
        'Dim formatRange As Excel.Range
        'Dim rowIndex As Integer = 1

        'Try
        '    wBook = _excel.Workbooks.Add()
        '    wSheet = wBook.ActiveSheet()

        '    formatRange = wSheet.Range("a1", "g" & TBLEmp.Count + 1)
        '    formatRange.Select()
        '    formatRange.NumberFormat = "@"

        '    wSheet.Cells(1, 1) = "NIK"
        '    wSheet.Cells(1, 2) = "Employee Code"
        '    wSheet.Cells(1, 3) = "Employee Name"
        '    wSheet.Cells(1, 4) = "Department Code"
        '    wSheet.Cells(1, 5) = "Department Description"
        '    wSheet.Cells(1, 6) = "Job Position Code"
        '    wSheet.Cells(1, 7) = "Job Position Description"
        '    wSheet.Cells(1, 8) = "Amount"

        '    TBLEmp.DefaultView.RowFilter = Nothing
        '    TBLEmp.DefaultView.Sort = "[Urutan] ASC"
        '    For i As Integer = 0 To TBLEmp.DefaultView.Count - 1
        '        RowEmp = TBLEmp.DefaultView(i).Row
        '        rowIndex += 1
        '        wSheet.Cells(rowIndex, 1) = RowEmp.NIK
        '        wSheet.Cells(rowIndex, 2) = RowEmp.Code
        '        wSheet.Cells(rowIndex, 3) = RowEmp.Full_Name
        '        wSheet.Cells(rowIndex, 4) = RowEmp.Dept_Code
        '        wSheet.Cells(rowIndex, 5) = RowEmp.Dept_Desc
        '        wSheet.Cells(rowIndex, 6) = RowEmp.Job_Code
        '        wSheet.Cells(rowIndex, 7) = RowEmp.Job_Desc
        '        wSheet.Cells(rowIndex, 8) = RowEmp.Amount
        '    Next

        '    wSheet.Columns.AutoFit()
        '    wBook.SaveAs(Me.FrmDia.FileName)

        '    Me.ReleaseObject(wSheet)
        '    wBook.Close(False)
        '    Me.ReleaseObject(wBook)
        '    _excel.Quit()
        '    Me.ReleaseObject(_excel)
        '    GC.Collect()

        '    MsgBox("File Export Success", MsgBoxStyle.Information, "Information")
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    _excel = Nothing
        '    wBook = Nothing
        '    wSheet = Nothing
        'End Try
    End Sub

    Private Sub ReleaseObject(ByVal o As Object)
        Dim count As Integer = System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
                count = count
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            o = Nothing
        End Try
    End Sub 

    Private Sub CopyFromSalaryCalculationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyFromSalaryCalculationToolStripMenuItem.Click
        Dim ij As New FrmTrans_Gaji_Override_PickDate
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary_item2TableAdapter
        Dim TBL As New DseTransaksiSalary.tbltrans_salary_item2DataTable
        Dim row As DseTransaksiSalary.tbltrans_salary_item2Row
        Dim OverrideReason As String = ""
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtFormulaCode.Text.Trim = "" Then
                MsgBox("Please specify salary component first.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            ij.IdentifierSalaryComponent = Me.TxtIdentifierSalaryFormula.Text
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.AllowUserToAddRows = False
                Me.dgvItem.Rows.Clear()

                OverrideReason = ij.TxtRemarksItem.Text.Trim
                DTA.Fill(TBL, Me.TxtIdentifierSalaryFormula.Text, ij.dtpDate.Value.Date)
                For Each row In TBL.Rows
                    Me.dgvItem.Rows.Add()

                    Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, Me.dgvItem.Rows.Count - 1).Value = row.Primary_Employee
                    Me.dgvItem.Item(Me.ColEmployeeCode.Index, Me.dgvItem.Rows.Count - 1).Value = row.Code
                    Me.dgvItem.Item(Me.ColEmployeeName.Index, Me.dgvItem.Rows.Count - 1).Value = row.Full_Name
                    Me.dgvItem.Item(Me.ColJobPositionCode.Index, Me.dgvItem.Rows.Count - 1).Value = row.Job_Position_Code
                    Me.dgvItem.Item(Me.ColJobPositionDescription.Index, Me.dgvItem.Rows.Count - 1).Value = row.Job_Position_Desc
                    Me.dgvItem.Item(Me.ColDepartmentCode.Index, Me.dgvItem.Rows.Count - 1).Value = row.Dept_Code
                    Me.dgvItem.Item(Me.ColDepartmentDesc.Index, Me.dgvItem.Rows.Count - 1).Value = row.Dept_Desc
                    Me.dgvItem.Item(Me.ColOverrideAmount.Index, Me.dgvItem.Rows.Count - 1).Value = Me.ObjCD.DigitGrouping(row.Amount, Me.Def_Prec)
                    Me.dgvItem.Item(Me.ColOverrideReason.Index, Me.dgvItem.Rows.Count - 1).Value = OverrideReason
                    Me.dgvItem.Item(Me.ColNo.Index, Me.dgvItem.Rows.Count - 1).Value = Me.dgvItem.RowCount
                Next

                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvItem.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub
End Class