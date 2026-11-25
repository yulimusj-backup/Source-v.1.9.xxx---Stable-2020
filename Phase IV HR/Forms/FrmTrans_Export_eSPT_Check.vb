Public Class FrmTrans_Export_eSPT_Check
    Public mParentForm1 As FrmTrans_Export_eSPT
    Public UserCode As String = ""

    Dim CheckPeriod As Date 
    Dim GrossIncome_DataCHeckStr = "Check the amount of gross income calculated every month = gross income calculated at the end of the year" + vbCrLf + _
    "(We ignore < IDR 100 difference for each employee due to rounding on gross income calculation)"
    Dim PrevIncome_DataCheckStr = "This check ensures that any previous income data is referenced by salary calculations"
    Dim BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCD As New Core_Module.ClsCekData 

    Private Sub FrmTrans_Export_eSPT_Check_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CheckPeriod = Me.mParentForm1.DtpPeriod.Value.Date
            Me.Text = Me.Text + " - Check period: " + Format(Me.CheckPeriod, "yyyy")
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboCheckType.SelectedIndex = 0 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCheckType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCheckType.SelectedIndexChanged
        Try
            Me.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.Rows.Clear()

            Me.ColGrossIncomeYearly.Visible = False
            Me.ColGrossIncomeMonthly.Visible = False
            Me.ColIncomeTaxMonthly.Visible = False
            Me.ColIncomeTaxYearly.Visible = False
            Me.ColA1FormNo.Visible = False
            Me.ColEmployer.Visible = False
            Me.ChkShowWrongDataOnly.Visible = False
            Me.ChkShowWrongDataOnly.Checked = False
            Me.ChkShowWrongDataOnly.TabStop = False

            If Me.cboCheckType.SelectedIndex = -1 Then
                Me.TxtCheckDesc.Text = "'" 
            ElseIf Me.cboCheckType.SelectedIndex = 0 Then
                Me.TxtCheckDesc.Text = Me.GrossIncome_DataCHeckStr
                Me.ColGrossIncomeYearly.Visible = True
                Me.ColGrossIncomeMonthly.Visible = True
                Me.ColIncomeTaxMonthly.Visible = True
                Me.ColIncomeTaxYearly.Visible = True
                Me.ChkShowWrongDataOnly.Visible = True
                Me.ChkShowWrongDataOnly.TabStop = True
            ElseIf Me.cboCheckType.SelectedIndex = 1 Then
                Me.TxtCheckDesc.Text = Me.PrevIncome_DataCheckStr
                Me.ColA1FormNo.Visible = True
                Me.ColEmployer.Visible = True
            End If

            Me.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Private Sub GrossIncomeCheck()
        Dim DGVParent As Windows.Forms.DataGridView
        Dim dtaMonthly As New DseTrans_Export_eSPTTableAdapters.tbltrans_salary_monthly1TableAdapter
        Dim tblMonthly As New DseTrans_Export_eSPT.tbltrans_salary_monthly1DataTable
        Dim rowMonthly As DseTrans_Export_eSPT.tbltrans_salary_monthly1Row
        Dim StartDate, EndDate As Date
        Dim rowDGV As Windows.Forms.DataGridViewRow
        Dim PrimaryEmployee As Double = -1
        Dim Ketemu As Boolean = False
        Dim Idx As Double = 0
        Dim PenghasilanBruto, PPH21 As Decimal

        Try
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor

            Me.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.Rows.Clear()

            dtaMonthly.Connection.ConnectionString = Core_Module.GetConstring
            StartDate = Me.ObjCD.GetFisrtDateOfTheYear(Me.CheckPeriod)
            EndDate = Me.ObjCD.GetLastDateOfTheYear(Me.CheckPeriod)

            'load data gross income dari data bulanan
            dtaMonthly.Fill(tblMonthly, EndDate, StartDate)
            tblMonthly.DefaultView.Sort = "[Primary Employee]"

            'data tahunan di ambil dari dgv parent form
            DGVParent = Me.mParentForm1.dgvA1
            For Each rowDGV In DGVParent.Rows
                PrimaryEmployee = rowDGV.Cells(Me.mParentForm1.ColA1_PrimaryEmployee.Index).Value
                If PrimaryEmployee <> -1 Then 
                    Ketemu = False
                    PenghasilanBruto = CDec(rowDGV.Cells(Me.mParentForm1.ColA1_Jumlah8.Index).Value)
                    PPH21 = CDec(rowDGV.Cells(Me.mParentForm1.ColA1_Jumlah20.Index).Value)
                    Idx = 0

                    While Idx <= tblMonthly.DefaultView.Count - 1
                        rowMonthly = tblMonthly.DefaultView(Idx).Row

                        If rowMonthly.Primary_Employee < PrimaryEmployee Then
                            Idx += 1
                        ElseIf rowMonthly.Primary_Employee = PrimaryEmployee Then
                            If Math.Abs(rowMonthly.Penghasilan_Bruto + rowMonthly._18PPh21YgTelahDipotongMasaPjkSblmnya1 - PenghasilanBruto) > 100 Or rowMonthly.Income_Tax - rowMonthly._21_PPH_Final_Pesangon <> PPH21 Then
                                Me.TambahBarisKeDGV(rowMonthly.Primary_Employee, rowMonthly.Code, rowMonthly.Full_Name, rowMonthly.Penghasilan_Bruto, _
                                PenghasilanBruto, rowMonthly.Income_Tax - rowMonthly._21_PPH_Final_Pesangon, PPH21, "", "", rowMonthly._18PPh21YgTelahDipotongMasaPjkSblmnya1)
                                If Me.ChkShowWrongDataOnly.Checked = False Then
                                    Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).DefaultCellStyle.BackColor = Drawing.Color.Red
                                End If
                            Else
                                If Me.ChkShowWrongDataOnly.Checked = False Then
                                    Me.TambahBarisKeDGV(rowMonthly.Primary_Employee, rowMonthly.Code, rowMonthly.Full_Name, rowMonthly.Penghasilan_Bruto, _
                                    PenghasilanBruto, rowMonthly.Income_Tax - rowMonthly._21_PPH_Final_Pesangon, PPH21, "", "", rowMonthly._18PPh21YgTelahDipotongMasaPjkSblmnya1)
                                End If
                            End If

                            Ketemu = True
                            tblMonthly.Rows.Remove(rowMonthly)
                            Exit While
                        Else
                            Exit While
                        End If
                    End While

                    If Ketemu = False Then
                        'tambahkan baris ke dgv
                        Me.TambahBarisKeDGV(rowDGV.Cells(Me.mParentForm1.ColA1_PrimaryEmployee.Index).Value, rowDGV.Cells(Me.mParentForm1.ColA1_EmployeeCode.Index).Value, rowDGV.Cells(Me.mParentForm1.ColA1_Nama.Index).Value, 0, _
                        PenghasilanBruto, 0, PPH21, "", "", 0)
                        If Me.ChkShowWrongDataOnly.Checked = False Then
                            Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).DefaultCellStyle.BackColor = Drawing.Color.Red
                        End If
                    End If

                End If 
            Next

            If tblMonthly.Rows.Count > 0 Then
                'tambahkan baris ke dgv
                For Each rowMonthly In tblMonthly.Rows
                    Me.TambahBarisKeDGV(rowMonthly.Primary_Employee, rowMonthly.Code, rowMonthly.Full_Name, rowMonthly.Penghasilan_Bruto, _
                    0, rowMonthly.Income_Tax - rowMonthly._21_PPH_Final_Pesangon, 0, "", "", rowMonthly._18PPh21YgTelahDipotongMasaPjkSblmnya1)
                    If Me.ChkShowWrongDataOnly.Checked = False Then
                        Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).DefaultCellStyle.BackColor = Drawing.Color.Red
                    End If
                Next
            End If

            Me.dgvItem.Sort(Me.ColEmployeeCode, ComponentModel.ListSortDirection.Ascending)
            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaMonthly.Dispose()
            tblMonthly.Dispose()
        End Try
    End Sub

    Private Sub TambahBarisKeDGV(ByVal vPrimaryEmployee As Double, ByVal vCode As String, ByVal vName As String, ByVal vGrossIncomeMonthly As Decimal, ByVal vGrossIncomeYearly As Decimal, _
    ByVal vPPH21Monthly As Decimal, ByVal vPPH21Yearly As Decimal, ByVal vNoFormA1 As String, ByVal vEmployer As String, ByVal vPrevPeriodIncomeTax As Decimal)
        Dim rowDGV As Windows.Forms.DataGridViewRow

        Try
            Me.dgvItem.Rows.Add()
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
            rowDGV.Cells(Me.ColEmployeeCode.Index).Value = vCode
            rowDGV.Cells(Me.ColEmployeeName.Index).Value = vName
            rowDGV.Cells(Me.ColA1FormNo.Index).Value = vNoFormA1
            rowDGV.Cells(Me.ColEmployer.Index).Value = vEmployer
            rowDGV.Cells(Me.ColGrossIncomeMonthly.Index).Value = Me.ObjCD.DigitGrouping(vGrossIncomeMonthly, 0)
            rowDGV.Cells(Me.ColGrossIncomeYearly.Index).Value = Me.ObjCD.DigitGrouping(vGrossIncomeYearly, 0)
            rowDGV.Cells(Me.ColIncomeTaxMonthly.Index).Value = Me.ObjCD.DigitGrouping(vPPH21Monthly, 0)
            rowDGV.Cells(Me.ColIncomeTaxYearly.Index).Value = Me.ObjCD.DigitGrouping(vPPH21Yearly, 0)
            rowDGV.Cells(Me.ColPrimaryEmployee.Index).Value = vPrimaryEmployee
            rowDGV.Cells(Me.ColGrossIncome_PrevIncomeTax.Index).Value = Me.ObjCD.DigitGrouping(vPrevPeriodIncomeTax, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PrevIncomeCheck()
        Try
            'load data prev. income pada tahun ybs

            'load data item prev. income pada item salary pada tahun ybs

            'pastikan datanya digunakan semua

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim ij As New FrmTrans_Export_eSPT_Check1
        Dim DTA As New DseTrans_Export_eSPTTableAdapters.tbltrans_salary_monthly2TableAdapter
        Dim TBL As New DseTrans_Export_eSPT.tbltrans_salary_monthly2DataTable
        Dim StartDate, EndDate As Date

        Try
            If e.ColumnIndex = Me.ColIncomeTaxMonthly.Index Or _
            e.ColumnIndex = Me.ColGrossIncome_PrevIncomeTax.Index Or _
            e.ColumnIndex = Me.ColGrossIncomeMonthly.Index Then
                StartDate = Me.ObjCD.GetFisrtDateOfTheYear(Me.CheckPeriod)
                EndDate = Me.ObjCD.GetLastDateOfTheYear(Me.CheckPeriod)
                DTA.Connection.ConnectionString = Core_Module.GetConstring

                DTA.Fill(TBL, EndDate, StartDate, Me.dgvItem.Item(Me.ColPrimaryEmployee.Index, e.RowIndex).Value)
                ij.TblRef = TBL

                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Try
            If Me.cboCheckType.SelectedIndex = 0 Then
                Me.GrossIncomeCheck()
            ElseIf Me.cboCheckType.SelectedIndex = 1 Then
                Me.PrevIncomeCheck() 
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class