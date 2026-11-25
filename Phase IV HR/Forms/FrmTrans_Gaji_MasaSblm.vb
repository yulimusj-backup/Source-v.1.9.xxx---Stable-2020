Option Explicit On
Option Strict Off

Imports system.Windows.Forms
Public Class FrmTrans_Gaji_MasaSblm
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Public vPrimaryLoad As Double = -1

    Dim objEmp As New Core_Module.ClsMasterKaryawan
    'Dim ObjAntrian As New Core_Module.ClsAntrian
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim Obj As New ClsTrans_Gaji_MasaSblm
    Dim ObjLicense As New ClsLicense

    Public Def_Prec As Double = Me.ObjCurr.GetDecimal_DefaultCurrency
    Dim baru As Boolean

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimary.BackColor, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.TxtNettIncome.Text = Me.ObjCD.DigitGrouping(Me.TxtNettIncome.Text, Me.Def_Prec)
            Me.TxtIncomeTaxPaid.Text = Me.ObjCD.DigitGrouping(Me.TxtIncomeTaxPaid.Text, Me.Def_Prec)

            Me.SplitContainer1.TabStop = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, False, False)
            Core_Module.ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)

            Me.CancelButton = Me.btnExit
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)

            Me.SplitContainer1.TabStop = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            'Core_Module.ModCekData.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            'Core_Module.ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, False, False)
            'Core_Module.ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, False, False)

            Me.TxtPrimary.Text = "-1"
            Me.TxtPrimaryEmployee.Text = "-1"
            Me.TxtA1FormNo.Text = ""
            Me.TxtEmployeeCode.Text = ""
            Me.TxtEmployeeName.Text = ""
            Me.TxtNettIncome.Text = 0
            Me.TxtIncomeTaxPaid.Text = 0
            Me.TxtEmployer.Text = ""
            Me.TxtRemarks.Text = ""

            Me.dtpDate.Value = Now
            Me.DtpWorkPeriodStart.Value = Me.ObjCD.GetFirstDateofTheMonth(Now)
            Me.DtpWorkPeriodTo.Value = Me.ObjCD.GetFirstDateofTheMonth(Now)

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim Dummy As String = ""

        Try
            'tidak boleh kosong
            KoleksiNamaData.Add("Employee Code")
            KoleksiTextBox.Add(Me.TxtEmployeeCode)
            KoleksiNamaData.Add("A1 Form No")
            KoleksiTextBox.Add(Me.TxtA1FormNo)
            KoleksiNamaData.Add("Employer")
            KoleksiTextBox.Add(Me.TxtEmployer)

            If Core_Module.ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'tidak boleh double
            If Me.baru = True Then
                If Me.Obj.ApakahSudahAda(Me.TxtA1FormNo.Text) = True Then
                    MsgBox("Duplication in a1 form no, please change your code", MsgBoxStyle.Information, "Information")
                    Me.TxtA1FormNo.Focus()
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.Obj.GetNoTransaksiByPrimary(Me.TxtPrimary.Text)
                If CodeLama <> Me.TxtA1FormNo.Text Then
                    If Me.Obj.ApakahSudahAda(Me.TxtA1FormNo.Text) = True Then
                        MsgBox("Duplication in a1 form no, please change your code", MsgBoxStyle.Information, "Information")
                        Me.TxtA1FormNo.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.Obj.GetTimeStampByPrimary(Me.TxtPrimary.Text)
                If Me.Obj.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.Obj.CariDetail(Me.TxtPrimary.Text, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If

                'Jika sudah ada transaksi salary atas karyawan tersebut yang menggunakan doc tsb, maka tidak bisa hapus, tidak bisa edit, tidak bisa simpan perubahan
                Dummy = Me.Obj.ApakahDocumentSudahDigunakanTransaksiSalary1(Me.TxtPrimary.Text)
                If Dummy <> "" Then
                    MsgBox("Previous income data is reference by salary calculation: " + Dummy + ", you can't save this data.", MsgBoxStyle.Information, "Information") 'msgbox english
                    Me.TxtEmployeeCode.Focus()
                    Return False
                End If
            End If

            'Date, kombinasi bulan tahun minimal sama dengan End Period, tidak boleh lebih kecil
            If Me.dtpDate.Value < Me.ObjCD.GetFirstDateofTheMonth(Me.DtpWorkPeriodTo.Value) Then
                MsgBox("Transaction date cannot be smaller than end period.", MsgBoxStyle.Information, "Information") 'msgbox english
                Me.dtpDate.Focus()
                Return False
            End If

            'Period start <= Period End
            If Me.ObjCD.GetFirstDateofTheMonth(Me.DtpWorkPeriodStart.Value) > Me.ObjCD.GetFirstDateofTheMonth(Me.DtpWorkPeriodTo.Value) Then
                MsgBox("Start period cannot be smaller than end period", MsgBoxStyle.Information, "Information") 'msgbox english
                Me.DtpWorkPeriodTo.Focus()
                Return False
            End If

            'Period start dan period end, harus di tahun yang sama
            If Me.DtpWorkPeriodStart.Value.Date.Year <> Me.DtpWorkPeriodTo.Value.Date.Year Then
                MsgBox("Start period and end period must be in the same year.", MsgBoxStyle.Information, "Information") 'msgbox english
                Me.DtpWorkPeriodTo.Focus()
                Return False
            End If

            'Period start dan period end, range maksimal 11 bulan
            If DateDiff(DateInterval.Month, Me.ObjCD.GetFirstDateofTheMonth(Me.DtpWorkPeriodStart.Value), _
            Me.ObjCD.GetFirstDateofTheMonth(Me.DtpWorkPeriodTo.Value)) >= 11 Then
                MsgBox("Start period and end period maximum range of 11 months.", MsgBoxStyle.Information, "Information") 'msgbox english
                Me.DtpWorkPeriodTo.Focus()
                Return False
            End If

            'satu employee, tidak bisa memiliki lebih dari 1 form A1 yang bersinggungan start period dan end period-nya
            If Me.Obj.ApakahEmployeeSudahAdaInPeriod(Me.TxtPrimaryEmployee.Text, Me.ObjCD.GetFirstDateofTheMonth(Me.DtpWorkPeriodStart.Value), _
            Me.ObjCD.GetFirstDateofTheMonth(Me.DtpWorkPeriodTo.Value), Me.TxtPrimary.Text) = True Then
                Me.TxtEmployeeCode.Focus()
                Return False
            End If

            'Jika sudah ada transaksi salary yang resign atau akhir tahun atas karyawan tersebut, maka tidak bisa simpan baru.
            'Jika sudah ada transaksi salary atas karyawan ybs yg <= end period, maka tidak bisa simpan/simpan perubahan
            'Dim dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salary_item_prev_incomeTableAdapter
            'Dim tbl As New DseTrans_Gaji_MasaSblm.tbltrans_salary_item_prev_incomeDataTable
            'Dim row As DseTrans_Gaji_MasaSblm.tbltrans_salary_item_prev_incomeRow
            Dim dta As New DseTrans_Gaji_MasaSblmTableAdapters.tbltrans_salaryTableAdapter
            Dim tbl As New DseTrans_Gaji_MasaSblm.tbltrans_salaryDataTable
            Dim row As DseTrans_Gaji_MasaSblm.tbltrans_salaryRow

            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(tbl, Me.TxtPrimaryEmployee.Text) > 0 Then
                For ij As Integer = 0 To tbl.Rows.Count - 1
                    row = tbl.Rows(ij)
                    With row
                        'Jika sudah ada transaksi salary yang resign maka tidak bisa simpan
                        'If .salary_is_resign <> 0 Then
                        '    MsgBox("This employee has a resignation salary calculation transaction, you can't save this data.", MsgBoxStyle.Information, "Information") 'msgbox english
                        '    Return False
                        'End If
                        ''Jika sudah ada transaksi salary atas karyawan ybs yg <= end period (atau akhir tahun atas karyawan tersebut) maka tidak bisa simpan
                        'If .salary_end_period <= Me.DtpWorkPeriodTo.Value.Date Then
                        '    MsgBox("This employee has a salary calculation transaction period <= " & Format(Me.DtpWorkPeriodTo.Value, "MMM yyyy") & ", you can't save this data.", MsgBoxStyle.Information, "Information") 'msgbox english
                        '    Return False
                        'End If
                    End With
                Next
            End If

            If Me.ObjCD.CekBilanganPositif(Me.TxtNettIncome.Text) = False Then
                MsgBox("Please enter positive number for nett income.", MsgBoxStyle.Information, "Information") 'msgbox english
                Me.TxtNettIncome.Focus()
                Return False
            End If

            'If Me.ObjCD.CekBilanganPositif(Me.TxtIncomeTaxPaid.Text) = False Then
            '    MsgBox("Please enter positive number for income tax paid.", MsgBoxStyle.Information, "Information") 'msgbox english
            '    Me.TxtIncomeTaxPaid.Focus()
            '    Return False
            'End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                .PrimaryEmployee = Me.TxtPrimaryEmployee.Text
                .NoTransaksi = Me.TxtA1FormNo.Text
                .Date_ = Me.dtpDate.Value
                .PeriodStart = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpWorkPeriodStart.Value)
                .PeriodEnd = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpWorkPeriodTo.Value)
                .Employer = Me.TxtEmployer.Text
                .NettIncome = Me.TxtNettIncome.Text
                .IncomeTaxPaid = Me.TxtIncomeTaxPaid.Text
                .Remarks = Me.TxtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtPrimaryEmployee.Text = .PrimaryEmployee
                Me.TxtEmployeeCode.Text = Me.objEmp.GetCodeByPrimary(.PrimaryEmployee)
                Me.TxtEmployeeName.Text = Me.objEmp.GetFullNameByPrimary(.PrimaryEmployee)
                Me.TxtA1FormNo.Text = .NoTransaksi
                Me.TxtEmployer.Text = .Employer
                Me.dtpDate.Value = .Date_
                Me.DtpWorkPeriodStart.Value = .PeriodStart
                Me.DtpWorkPeriodTo.Value = .PeriodEnd
                Me.TxtNettIncome.Text = Me.ObjCD.DigitGrouping(.NettIncome, Me.Def_Prec)
                Me.TxtIncomeTaxPaid.Text = Me.ObjCD.DigitGrouping(.IncomeTaxPaid, Me.Def_Prec)
                Me.TxtRemarks.Text = .Remarks

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_MasaSblm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            If Me.TxtRemarks.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            If Me.Obj.SimpanBaru(Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me) = True Then
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
                ElseIf hasil = Windows.Forms.DialogResult.No Then
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_MasaSblm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            Me.DtpWorkPeriodStart.Value = Me.ObjCD.GetFirstDateofTheMonth(Now)
            Me.DtpWorkPeriodTo.Value = Me.ObjCD.GetFirstDateofTheMonth(Now)

            'Me.txtCode.Width = Me.SplitContainer1.Panel1.Width - Me.txtCode.Left
            'Me.TxtDesc.Width = Me.SplitContainer1.Panel1.Width - Me.TxtDesc.Left
            'Me.TxtRemarks.Width = Me.SplitContainer1.Panel2.Width - Me.TxtRemarks.Left

            If Me.vPrimaryLoad <> -1 Then
                Me.ViewData()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ViewData()
        Try
            If Me.Obj.CariDetail(Me.vPrimaryLoad, Me) = True Then
                Me.baru = False
                Me.SalinKeField()
                Me.btnEdit.Focus()
                Me.NonAktifkanDetail()
            Else
                Me.btnFind.Focus()
            End If
            Me.btnFind.Focus()

            Me.btnNew.Enabled = False
            Me.btnFind.Enabled = False
            Me.btnEdit.Enabled = False
            Me.btnDelete.Enabled = False
            Me.btnPrint.Enabled = False

            Me.MinimizeBox = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult

        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanBaru(Me.UserCode, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me) Then
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
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.TxtEmployeeCode.Text <> "" Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Me.NonAktifkanDetail()
                        Me.Obj.CariDetail(Me.TxtPrimary.Text, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.NonAktifkanDetail()
                Me.KosongkanField()
            End If

            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTrans_Gaji_MasaSblm_Cari

        Try
            hj.BackColorAll = Me.BackColorAll
            hj.Frm = Me
            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                If Me.Obj.CariDetail(hj.PrimaryTransaksi, Me) = True Then
                    Me.baru = False
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                    Me.NonAktifkanDetail()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True

            Me.KosongkanField()
            Me.AktifkanDetail()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.TxtEmployeeCode.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtEmployeeCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtEmployeeCode.Text <> "" Then
                If Me.Obj.ApakahBolehedit(Me.UserCode, Me.TxtPrimary.Text) = True Then
                    Me.baru = False
                    Me.Obj.CariDetail(Me.TxtPrimary.Text, Me)
                    Me.AktifkanDetail()
                    Me.TxtEmployeeCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.TxtEmployeeCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtEmployeeCode.Text <> "" Then
                If Me.Obj.ApakahBolehHapus(Me.UserCode, Me.TxtPrimary.Text) = True Then
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.Hapus(Me.TxtPrimary.Text, Me.UserCode, Me) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        'Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi
        Dim FrmCetak As New FrmTrans_Gaji_MasaSblm_Cetak

        Try
            'If Me.TxtPrimary.Text = "" Or Me.TxtPrimary.Text = -1 Then
            '    Me.btnFind.PerformClick()
            'End If

            'If Me.TxtPrimary.Text <> "" And Me.TxtPrimary.Text <> -1 Then
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            'If Me.Obj.Cetak(Me.TxtPrimary.Text, FrmCetak.crv, FrmCetak.crvDetail) = True Then
            FrmCetak.PrimaryTrans = Me.TxtPrimary.Text
            FrmCetak.ShowDialog()
            'End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'FrmO.Dispose()
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub TxtEmployeeCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtEmployeeCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.TxtEmployeeCode_Leave(sender, e)
                Me.TxtA1FormNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtEmployeeCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmployeeCode.Leave
        Try
            If Me.TxtEmployeeCode.Text <> "" Then
                If Me.objEmp.ApakahSudahAda(Me.TxtEmployeeCode.Text) = True Then
                    Me.TxtPrimaryEmployee.Text = Me.objEmp.GetPrimaryByCode(Me.TxtEmployeeCode.Text)
                    Me.TxtEmployeeCode.Text = Me.objEmp.GetCodeByPrimary(Me.TxtPrimaryEmployee.Text)
                    Me.TxtEmployeeName.Text = Me.objEmp.GetFullNameByPrimary(Me.TxtPrimaryEmployee.Text)
                Else
                    Me.btnFindEmployee.PerformClick()
                End If
            Else
                Me.TxtPrimaryEmployee.Text = -1
                Me.TxtEmployeeCode.Text = ""
                Me.TxtEmployeeName.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindEmployee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindEmployee.Click
        Dim vFrm As New Core_Module.FrmMasterKaryawanCari
        Try
            vFrm.txtCari.Text = Me.TxtEmployeeCode.Text
            vFrm.BackColorAll = Me.BackColorAll
            vFrm.ShowDialog()
            If vFrm.PrimaryEmploye > -1 Then
                Me.TxtPrimaryEmployee.Text = vFrm.PrimaryEmploye
                Me.TxtEmployeeCode.Text = Me.objEmp.GetCodeByPrimary(vFrm.PrimaryEmploye)
                Me.TxtEmployeeName.Text = Me.objEmp.GetFullNameByPrimary(vFrm.PrimaryEmploye)
            Else
                Me.TxtPrimaryEmployee.Text = -1
                Me.TxtEmployeeCode.Text = ""
                Me.TxtEmployeeName.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vFrm.Dispose()
        End Try
    End Sub

    Private Sub TxtNettIncome_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNettIncome.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtNettIncome, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNettIncome_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNettIncome.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtNettIncome, sender, e)
        Me.TxtNettIncome.MaxLength = 14
    End Sub

    Private Sub TxtNettIncome_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNettIncome.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtNettIncome, sender, e)
    End Sub

    Private Sub TxtNettIncome_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNettIncome.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtNettIncome, sender, e, Me.Def_Prec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtIncomeTaxPaid_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtIncomeTaxPaid.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtIncomeTaxPaid, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtIncomeTaxPaid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtIncomeTaxPaid.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtIncomeTaxPaid, sender, e)
        Me.TxtIncomeTaxPaid.MaxLength = 14
    End Sub

    Private Sub TxtIncomeTaxPaid_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtIncomeTaxPaid.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtIncomeTaxPaid, sender, e)
    End Sub

    Private Sub TxtIncomeTaxPaid_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtIncomeTaxPaid.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtIncomeTaxPaid, sender, e, Me.Def_Prec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtA1FormNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtA1FormNo.GotFocus, TxtEmployeeCode.GotFocus, TxtEmployer.GotFocus, TxtNettIncome.GotFocus, TxtIncomeTaxPaid.GotFocus, TxtRemarks.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtEmployeeName.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtA1FormNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtA1FormNo.LostFocus, TxtEmployeeCode.LostFocus, TxtEmployer.LostFocus, TxtNettIncome.LostFocus, TxtIncomeTaxPaid.LostFocus, TxtRemarks.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtEmployeeName.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimary.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class