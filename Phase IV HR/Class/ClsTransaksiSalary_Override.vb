Option Explicit On
Option Strict Off

Imports Core_Module, System.Windows.Forms, System.Drawing
Public Class ClsTransaksiSalary_Override
    Private mPrimary As Double = -1
    Private mDocumentNo As String = ""
    Private mTanggalTransaksi As Date = #1/1/1950#
    Private mFormulaType As EnumFormulaType = EnumFormulaType.Salary_Formula
    Private mPrimaryFormula As Double = -1
    Private mIdentifierFormula As String = ""
    Private mRemarks As String = ""
    Private mUrutan As Integer = 0
    Private mTipe As EnumType = EnumType.SalaryOverride

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum EnumFormulaType As Integer
        Salary_Formula
        Salary_Formula_Other_Deduction
    End Enum

    Public Enum EnumType As Integer
        SalaryOverride
        OtherPayment
    End Enum

    Public Property Tipe() As EnumType
        Get
            Return Me.mTipe
        End Get
        Set(ByVal value As EnumType)
            mTipe = value
        End Set
    End Property

    Public Property Urutan() As Integer
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Integer)
            mUrutan = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property DocumentNo() As String
        Get
            Return mDocumentNo
        End Get
        Set(ByVal value As String)
            mDocumentNo = value
        End Set
    End Property

    Public Property TanggalTransaksi() As Date
        Get
            Return mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property PrimaryFormula() As Double
        Get
            Return mPrimaryFormula
        End Get
        Set(ByVal value As Double)
            mPrimaryFormula = value
        End Set
    End Property

    Public Property IdentifierFormula() As String
        Get
            Return mIdentifierFormula
        End Get
        Set(ByVal value As String)
            mIdentifierFormula = value
        End Set
    End Property

    Public Property FormulaType() As EnumFormulaType
        Get
            Return mFormulaType
        End Get
        Set(ByVal value As EnumFormulaType)
            mFormulaType = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDocumentNo(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetDocumentNo(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetFormulaType(ByVal vPrimary As Double) As EnumFormulaType
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim dummy As System.Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dummy = DTA.ScalarQuery_GetFormulaType(vPrimary)
            If dummy.HasValue Then
                Return dummy
            Else
                Return EnumFormulaType.Salary_Formula
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetIdentifierSalaryFormula(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetIdentifierSalaryFormula(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return "-1-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-1-1"
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByDocumentNo(ByVal vDocumentNo As String) As Integer
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dummy = DTA.ScalarQuery_GetJmlByDocumentNo(vDocumentNo)
            If dummy.HasValue Then
                Return dummy
            Else
                Return EnumFormulaType.Salary_Formula
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vDocumentNo As String) As Double
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimary(vDocumentNo)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryFormula(Primary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTanggalTransaksi(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return "#1/1/1950#"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTimestamp(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return "#1/1/1950#"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByIdentifierFormula(ByVal vIdentifierFormula As String, ByVal vPrimaryUser As Double) As Double
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tblmaster_salary_formula_salarycomponent1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByIdentifierFormula(vIdentifierFormula, vPrimaryUser)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByComponentName(ByVal vComponentName As String, ByVal vPrimaryUser As Double) As Double
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tblmaster_salary_formula_salarycomponent1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByComponentName(vComponentName, vPrimaryUser)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda_IdentifierFormula(ByVal vIdentifierFormula As String, ByVal vPrimaryUser As Double) As Double
        Try
            If Me.GetJmlByIdentifierFormula(vIdentifierFormula, vPrimaryUser) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSudahAda_FormulaSalary(ByVal vComponentName As String, ByVal vPrimaryUser As Double) As Double
        Try
            If Me.GetJmlByComponentName(vComponentName, vPrimaryUser) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJml_Formula(ByVal vKriteria As String, ByVal vSearchBy As Core_Module.EnumSearchCriteria, ByVal vPrimaryUser As Double, _
     Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Double
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tblcarimaster_salary_formula_view1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByComponentName Then
                Return DTA.ScalarQuery_GetJmlByComponentNameLIKE(vKriteria + "%", vPrimaryUser).Value
            ElseIf vSearchBy = EnumSearchCriteria.BySalaryFormula Then
                Return DTA.ScalarQuery_GetJmlBySalaryFormulaLIKE(vKriteria + "%", vPrimaryUser).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV_Formula(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalary_Override.tblcarimaster_salary_formula_view1DataTable, _
    ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal vPrimaryUser As Double, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Boolean
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tblcarimaster_salary_formula_view1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            If vPrimaryUser = -1 Then
                If SearchCriteria = EnumSearchCriteria.ByComponentName Then
                    QueryResult = DTA.FillBy_ComponentNameLike2(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.BySalaryFormula Then
                    QueryResult = DTA.FillBy_SalaryFormulaLike2(TBL, Criteria + "%")
                End If
            Else
                If SearchCriteria = EnumSearchCriteria.ByComponentName Then
                    QueryResult = DTA.FillBy_ComponentNameLIKE(TBL, Criteria + "%", vPrimaryUser)
                ElseIf SearchCriteria = EnumSearchCriteria.BySalaryFormula Then
                    QueryResult = DTA.FillBy_SalaryFormulaLIKE(TBL, Criteria + "%", vPrimaryUser)
                End If
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(2).Visible = False
                DGV.Columns(3).Visible = False
                'DGV.Columns(0).Frozen = True
                'DGV.Columns(1).Frozen = True
                'DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                'DGV.Columns(1).HeaderText = "Code"
                'DGV.Columns(2).HeaderText = "Description"
                'DGV.Columns(3).HeaderText = "City"
                'DGV.Columns(4).HeaderText = "Remarks"
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > TimeSpan.TicksPerSecond Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, _
                    0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml_OtherDeductionFormula(ByVal vKriteria As String, ByVal vSearchBy As Core_Module.EnumSearchCriteria, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Double
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tblcarimaster_salary_formula_view2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByComponentName Then
                Return DTA.ScalarQuery_GetJmlByComponentNameLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.BySalaryFormula Then
                Return DTA.ScalarQuery_GetJmlBySalaryFormulaLIKE(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV_OtherDeductionFormula(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalary_Override.tblcarimaster_salary_formula_view2DataTable, _
    ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Boolean
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tblcarimaster_salary_formula_view2TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            If SearchCriteria = EnumSearchCriteria.ByComponentName Then
                QueryResult = DTA.FillBy_ComponentNameLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalaryFormula Then
                QueryResult = DTA.FillBy_SalaryFormulaLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                'DGV.Columns(0).Frozen = True
                'DGV.Columns(1).Frozen = True
                'DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                'DGV.Columns(1).HeaderText = "Code"
                'DGV.Columns(2).HeaderText = "Description"
                'DGV.Columns(3).HeaderText = "City"
                'DGV.Columns(4).HeaderText = "Remarks"
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > TimeSpan.TicksPerSecond Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, _
                    0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    'Fungsi ini copy daru master cuti
    Public Enum enumCariData
        dept
        Job
        Other
        AllEmp
    End Enum

    Private Function GetBS(ByVal vPrimaryKaryawan As Double, ByVal TBLBS As DseTransaksiSalary_Override.tblmaster_salary_formula_component_detailemployeeDataTable) As Decimal
        Dim rowBS As DseTransaksiSalary_Override.tblmaster_salary_formula_component_detailemployeeRow

        Try
            TBLBS.DefaultView.RowFilter = "[Primary Karyawan] = " + vPrimaryKaryawan.ToString
            If TBLBS.DefaultView.Count = 0 Then
                Return 0
            Else
                rowBS = TBLBS.DefaultView(0).Row
                Return rowBS.Basic_Salary
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub CariDataKaryawan(ByVal vCariBerdasakan As enumCariData, ByVal vPrimary As Double, _
    ByVal vDefPrec As Double, ByVal frm As FrmTrans_Gaji_Override)
        'Tabel untuk Salary Formula
        Dim dta As New DseTransaksiSalary_OverrideTableAdapters.DataTable2TableAdapter
        Dim tbl As New DseTransaksiSalary_Override.DataTable2DataTable
        Dim row As DseTransaksiSalary_Override.DataTable2Row
        Dim dtaItem As New DseTransaksiSalary_OverrideTableAdapters.DataTable3TableAdapter
        Dim tblItem As New DseTransaksiSalary_Override.DataTable3DataTable
        Dim rowItem As DseTransaksiSalary_Override.DataTable3Row
        'Tabel untuk Other Deduction
        Dim dta1 As New DseTransaksiSalary_OverrideTableAdapters.DataTable4TableAdapter
        Dim tbl1 As New DseTransaksiSalary_Override.DataTable4DataTable
        Dim row1 As DseTransaksiSalary_Override.DataTable4Row
        Dim dtaItem1 As New DseTransaksiSalary_OverrideTableAdapters.DataTable5TableAdapter
        Dim tblItem1 As New DseTransaksiSalary_Override.DataTable5DataTable
        Dim rowItem1 As DseTransaksiSalary_Override.DataTable5Row
        '----------------------------
        Dim rowdgv As DataGridViewRow
        Dim FrmDialog As New HR.FrmAddReplace
        Dim masukan As Boolean = False
        Dim ObjKaryawan As New ClsMasterKaryawan
        Dim ObjCD As New Core_Module.ClsCekData
        Dim mAmount As Decimal = 0
        Dim mReason As String = ""
        Dim ObjUser As New Core_Module.ClsUser
        Dim mPrimaryUser As Double = ObjUser.GetPrimary(frm.UserCode)
        Dim mTypePencarian As EnumFormulaType = EnumFormulaType.Salary_Formula
        Dim ByBS As Boolean = False
        Dim DTABS As New DseTransaksiSalary_OverrideTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Dim TBLBS As New DseTransaksiSalary_Override.tblmaster_salary_formula_component_detailemployeeDataTable
        Dim FaktorKali As Decimal = 0

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'connection string 
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dtaItem.Connection.ConnectionString = Core_Module.GetConstring
            dta1.Connection.ConnectionString = Core_Module.GetConstring
            dtaItem1.Connection.ConnectionString = Core_Module.GetConstring
            DTABS.Connection.ConnectionString = dta.Connection.ConnectionString

            'Tentukan Type Pencarian Salary Formula ato Others
            If frm.CboFormulaType.SelectedIndex = 0 Then
                mTypePencarian = EnumFormulaType.Salary_Formula
            Else
                mTypePencarian = EnumFormulaType.Salary_Formula_Other_Deduction
            End If

            If mTypePencarian = EnumFormulaType.Salary_Formula Then
                'Salary Formula 
                ''filter data
                If vCariBerdasakan = enumCariData.Job Then
                    dta.FillBy_PrimaryJob(tbl, vPrimary, mPrimaryUser, frm.TxtPrimarySalaryFormula.Text)
                ElseIf vCariBerdasakan = enumCariData.dept Then
                    dta.FillBy_PrimaryDept(tbl, vPrimary, mPrimaryUser, frm.TxtPrimarySalaryFormula.Text)
                ElseIf vCariBerdasakan = enumCariData.Other Then
                    dtaItem.FillBy(tblItem, mPrimaryUser, vPrimary)
                ElseIf vCariBerdasakan = enumCariData.AllEmp Then
                    dta.Fill(tbl, mPrimaryUser, frm.TxtPrimarySalaryFormula.Text)
                End If

                ''pilih add/replace
                'ok = Add
                'cancel = Replace()
                FrmDialog.ShowDialog()
                FrmDialog.BackColorAll = frm.BackColorAll
                If FrmDialog.DialogResult = DialogResult.OK Then
                    masukan = False
                ElseIf FrmDialog.DialogResult = DialogResult.Cancel Then
                    frm.dgvItem.Rows.Clear()
                    masukan = True
                ElseIf FrmDialog.DialogResult = DialogResult.No Then
                    Exit Sub
                End If

                mAmount = FrmDialog.TxtOverrideAmount.Text
                mReason = FrmDialog.TxtRemarksItem.Text
                FaktorKali = FrmDialog.TxtBS.Text
                If FrmDialog.rboByBS.Checked = True Then
                    ByBS = True
                    DTABS.Fill(TBLBS)
                Else
                    ByBS = False
                End If

                ''pindahkan ke dgv
                If vCariBerdasakan = enumCariData.dept Or vCariBerdasakan = enumCariData.Job Or vCariBerdasakan = enumCariData.AllEmp Then
                    'FrmDialog.ShowDialog()
                    'FrmDialog.BackColorAll = frm.BackColorAll
                    'If FrmDialog.DialogResult = DialogResult.OK Then
                    '    masukan = False
                    'ElseIf FrmDialog.DialogResult = DialogResult.Cancel Then
                    '    frm.dgvItem.Rows.Clear()
                    '    masukan = True
                    'ElseIf FrmDialog.DialogResult = DialogResult.No Then
                    '    Exit Sub
                    'End If

                    'berdasarkan departemen atau job posisi
                    For Each row In tbl.Rows
                        'cek data apakah sudah ada di dgv
                        'jika replace lewati pengecekan
                        If FrmDialog.DialogResult = DialogResult.OK Then
                            For Each rowdgv In frm.dgvItem.Rows
                                If rowdgv.Cells(frm.ColPrimaryEmployee.Index).Value <> row.Primary Then
                                    masukan = True
                                Else
                                    masukan = False
                                    Exit For
                                End If
                            Next
                        End If
                        If masukan = True Then
                            frm.dgvItem.Rows.Add()
                            If frm.dgvItem.Rows.Count > 1 Then
                                rowdgv = frm.dgvItem.Rows(frm.dgvItem.RowCount - 2)
                            Else
                                rowdgv = frm.dgvItem.Rows(frm.dgvItem.RowCount)
                            End If
                            rowdgv.Cells(frm.ColEmployeeCode.Index).Value = row.Code
                            rowdgv.Cells(frm.ColEmployeeName.Index).Value = row.Full_Name
                            rowdgv.Cells(frm.ColDepartmentCode.Index).Value = row.Dept_Code
                            rowdgv.Cells(frm.ColDepartmentDesc.Index).Value = row.Dept_Desc
                            rowdgv.Cells(frm.ColJobPositionCode.Index).Value = row.Job_Code
                            rowdgv.Cells(frm.ColJobPositionDescription.Index).Value = row.Job_Desc
                            If ByBS = True Then
                                rowdgv.Cells(frm.ColOverrideAmount.Index).Value = ObjCD.DigitGrouping(Me.GetBS(row.Primary, TBLBS) * FaktorKali, vDefPrec)
                            Else
                                rowdgv.Cells(frm.ColOverrideAmount.Index).Value = ObjCD.DigitGrouping(mAmount, vDefPrec)
                            End If
                            rowdgv.Cells(frm.ColOverrideReason.Index).Value = mReason
                            rowdgv.Cells(frm.ColPrimaryEmployee.Index).Value = row.Primary
                        End If
                    Next
                ElseIf vCariBerdasakan = enumCariData.Other Then
                    'berdasarkan data cuti yang sudah ada
                    For Each rowItem In tblItem.Rows
                        'cek data apakah sudah ada di dgv
                        'jika replace lewati pengecekan
                        If FrmDialog.DialogResult = DialogResult.OK Then
                            For Each rowdgv In frm.dgvItem.Rows
                                If rowdgv.Cells(frm.ColPrimaryEmployee.Index).Value <> rowItem.Primary_Employee Then
                                    masukan = True
                                Else
                                    masukan = False
                                    Exit For
                                End If
                            Next
                        End If
                        If masukan = True Then
                            frm.dgvItem.Rows.Add()
                            If frm.dgvItem.Rows.Count > 1 Then
                                rowdgv = frm.dgvItem.Rows(frm.dgvItem.RowCount - 2)
                            Else
                                rowdgv = frm.dgvItem.Rows(frm.dgvItem.RowCount)
                            End If
                            rowdgv.Cells(frm.ColEmployeeCode.Index).Value = rowItem.Employee_Code
                            rowdgv.Cells(frm.ColEmployeeName.Index).Value = rowItem.Employee_Name
                            rowdgv.Cells(frm.ColDepartmentCode.Index).Value = rowItem.Dept_Code
                            rowdgv.Cells(frm.ColDepartmentDesc.Index).Value = rowItem.Dept_Desc
                            rowdgv.Cells(frm.ColJobPositionCode.Index).Value = rowItem.Job_Code
                            rowdgv.Cells(frm.ColJobPositionDescription.Index).Value = rowItem.Job_Desc
                            If ByBS = True Then
                                rowdgv.Cells(frm.ColOverrideAmount.Index).Value = ObjCD.DigitGrouping(Me.GetBS(row.Primary, TBLBS) * FaktorKali, vDefPrec)
                            Else
                                rowdgv.Cells(frm.ColOverrideAmount.Index).Value = ObjCD.DigitGrouping(mAmount, vDefPrec)
                            End If
                            rowdgv.Cells(frm.ColOverrideReason.Index).Value = mReason
                            rowdgv.Cells(frm.ColPrimaryEmployee.Index).Value = rowItem.Primary_Employee
                        End If
                    Next
                End If
                ObjCD.SetNumber(frm.dgvItem)
            Else
                'Other Deduction
                If vCariBerdasakan = enumCariData.Job Then
                    dta1.FillBy_PrimaryJob(tbl1, frm.TxtPrimarySalaryFormula.Text, vPrimary)
                ElseIf vCariBerdasakan = enumCariData.dept Then
                    dta1.FillBy_PrimaryDept(tbl1, frm.TxtPrimarySalaryFormula.Text, vPrimary)
                ElseIf vCariBerdasakan = enumCariData.Other Then
                    dtaItem1.Fill(tblItem1, frm.TxtPrimarySalaryFormula.Text, vPrimary)
                ElseIf vCariBerdasakan = enumCariData.AllEmp Then
                    dta1.Fill(tbl1, frm.TxtPrimarySalaryFormula.Text)
                End If

                ''pilih add/replace
                'ok = Add
                'cancel = replace
                FrmDialog.ShowDialog()
                FrmDialog.BackColorAll = frm.BackColorAll
                If FrmDialog.DialogResult = DialogResult.OK Then
                    masukan = False
                ElseIf FrmDialog.DialogResult = DialogResult.Cancel Then
                    frm.dgvItem.Rows.Clear()
                    masukan = True
                ElseIf FrmDialog.DialogResult = DialogResult.No Then
                    Exit Sub
                End If

                mAmount = FrmDialog.TxtOverrideAmount.Text
                mReason = FrmDialog.TxtRemarksItem.Text
                FaktorKali = FrmDialog.TxtBS.Text
                If FrmDialog.rboByBS.Checked = True Then
                    ByBS = True
                    DTABS.Fill(TBLBS)
                Else
                    ByBS = False
                End If

                ''pindahkan ke dgv
                If vCariBerdasakan = enumCariData.dept Or vCariBerdasakan = enumCariData.Job Or vCariBerdasakan = enumCariData.AllEmp Then
                    'berdasarkan departemen atau job posisi
                    For Each row1 In tbl1.Rows
                        'cek data apakah sudah ada di dgv
                        'jika replace lewati pengecekan
                        If FrmDialog.DialogResult = DialogResult.OK Then
                            For Each rowdgv In frm.dgvItem.Rows
                                If rowdgv.Cells(frm.ColPrimaryEmployee.Index).Value <> row1.Primary Then
                                    masukan = True
                                Else
                                    masukan = False
                                    Exit For
                                End If
                            Next
                        End If
                        If masukan = True Then
                            frm.dgvItem.Rows.Add()
                            If frm.dgvItem.Rows.Count > 1 Then
                                rowdgv = frm.dgvItem.Rows(frm.dgvItem.RowCount - 2)
                            Else
                                rowdgv = frm.dgvItem.Rows(frm.dgvItem.RowCount)
                            End If
                            rowdgv.Cells(frm.ColEmployeeCode.Index).Value = row1.Code
                            rowdgv.Cells(frm.ColEmployeeName.Index).Value = row1.Full_Name
                            rowdgv.Cells(frm.ColDepartmentCode.Index).Value = row1.Dept_Code
                            rowdgv.Cells(frm.ColDepartmentDesc.Index).Value = row1.Dept_Desc
                            rowdgv.Cells(frm.ColJobPositionCode.Index).Value = row1.Job_Code
                            rowdgv.Cells(frm.ColJobPositionDescription.Index).Value = row1.Job_Desc
                            If ByBS = True Then
                                rowdgv.Cells(frm.ColOverrideAmount.Index).Value = ObjCD.DigitGrouping(Me.GetBS(row.Primary, TBLBS) * FaktorKali, vDefPrec)
                            Else
                                rowdgv.Cells(frm.ColOverrideAmount.Index).Value = ObjCD.DigitGrouping(mAmount, vDefPrec)
                            End If
                            rowdgv.Cells(frm.ColOverrideReason.Index).Value = mReason
                            rowdgv.Cells(frm.ColPrimaryEmployee.Index).Value = row1.Primary
                        End If
                    Next
                ElseIf vCariBerdasakan = enumCariData.Other Then
                    'berdasarkan data cuti yang sudah ada
                    For Each rowItem1 In tblItem1.Rows
                        'cek data apakah sudah ada di dgv
                        'jika replace lewati pengecekan
                        If FrmDialog.DialogResult = DialogResult.OK Then
                            For Each rowdgv In frm.dgvItem.Rows
                                If rowdgv.Cells(frm.ColPrimaryEmployee.Index).Value <> rowItem1.Primary_Employee Then
                                    masukan = True
                                Else
                                    masukan = False
                                    Exit For
                                End If
                            Next
                        End If
                        If masukan = True Then
                            frm.dgvItem.Rows.Add()
                            If frm.dgvItem.Rows.Count > 1 Then
                                rowdgv = frm.dgvItem.Rows(frm.dgvItem.RowCount - 2)
                            Else
                                rowdgv = frm.dgvItem.Rows(frm.dgvItem.RowCount)
                            End If
                            rowdgv.Cells(frm.ColEmployeeCode.Index).Value = rowItem1.Employee_Code
                            rowdgv.Cells(frm.ColEmployeeName.Index).Value = rowItem1.Employee_Name
                            rowdgv.Cells(frm.ColDepartmentCode.Index).Value = rowItem1.Dept_Code
                            rowdgv.Cells(frm.ColDepartmentDesc.Index).Value = rowItem1.Dept_Desc
                            rowdgv.Cells(frm.ColJobPositionCode.Index).Value = rowItem1.Job_Code
                            rowdgv.Cells(frm.ColJobPositionDescription.Index).Value = rowItem1.Job_Desc
                            If ByBS = True Then
                                rowdgv.Cells(frm.ColOverrideAmount.Index).Value = ObjCD.DigitGrouping(Me.GetBS(row.Primary, TBLBS) * FaktorKali, vDefPrec)
                            Else
                                rowdgv.Cells(frm.ColOverrideAmount.Index).Value = ObjCD.DigitGrouping(mAmount, vDefPrec)
                            End If
                            rowdgv.Cells(frm.ColOverrideReason.Index).Value = mReason
                            rowdgv.Cells(frm.ColPrimaryEmployee.Index).Value = rowItem1.Primary
                        End If
                    Next
                End If
                ObjCD.SetNumber(frm.dgvItem)
            End If

            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            tbl.Dispose()
            dtaItem.Dispose()
            tblItem.Dispose()
            dta1.Dispose()
            tbl1.Dispose()
            dtaItem1.Dispose()
            tblItem1.Dispose()
            FrmDialog.Dispose()
            ObjKaryawan = Nothing
            ObjCD = Nothing
            DTABS.Dispose()
            TBLBS.Dispose()
        End Try
    End Sub

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vObjCD As Core_Module.ClsCekData, _
     ByVal vDGVItem As DataGridView, ByVal vDef_Prec As Double, ByVal i As FrmTrans_Gaji_Override) As Boolean
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim TBL As New DseTransaksiSalary_Override.tbltrans_salary_override_amount1DataTable
        Dim row As DseTransaksiSalary_Override.tbltrans_salary_override_amount1Row
        Dim ObjItem As New ClsTransaksiSalary_Override_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mDocumentNo = .Document_No
                    Me.mTanggalTransaksi = ._Date
                    Me.mUrutan = .Urutan
                    Me.mFormulaType = .Formula_Type
                    Me.mRemarks = .Remarks
                    Me.mPrimaryFormula = .Primary_Formula
                    Me.mIdentifierFormula = .Identifier_Formula
                    Me.mTipe = .tipe

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'cari data item
                ObjItem.IsiDGVItem(Primary, vDGVItem, vDef_Prec, i)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As Core_Module.EnumSearchCriteria, _
    ByVal vTipe As EnumType, Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Double
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tblcarimaster_salary_formula_view3TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByDocumentNoLIKE(vKriteria + "%", CInt(vTipe)).Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByDate Then
                Return DTA.ScalarQuery_GetJmlByTanggalDariSampai(vEndDate, vStartDate, CInt(vTipe)).Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%", CInt(vTipe)).Value
            ElseIf vSearchBy = EnumSearchCriteria.ByComponentName Then
                Return DTA.ScalarQuery_GetJmlByComponentName(vKriteria + "%", CInt(vTipe)).Value
            ElseIf vSearchBy = EnumSearchCriteria.BySalaryFormula Then
                Return DTA.ScalarQuery_GetJmlBySalaryFormulaLIKE(vKriteria + "%", CInt(vTipe)).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV2(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalary_Override.tblcarimaster_salary_formula_view3DataTable, _
    ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal vTipe As EnumType, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#, Optional ByVal vUserCode As String = "") As Boolean
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tblcarimaster_salary_formula_view3TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If vUserCode = "" Then
                If SearchCriteria = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                    QueryResult = DTA.FillBy_DocumentNoLIKE(TBL, Criteria + "%", CInt(vTipe))
                ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDate Then
                    QueryResult = DTA.FillBy_TanggalTransaksi(TBL, vEndDate, vStartDate, CInt(vTipe))
                ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%", CInt(vTipe))
                ElseIf SearchCriteria = EnumSearchCriteria.ByComponentName Then
                    QueryResult = DTA.FillBy_ComponentNameLIKE(TBL, Criteria + "%", CInt(vTipe))
                ElseIf SearchCriteria = EnumSearchCriteria.BySalaryFormula Then
                    QueryResult = DTA.FillBy_SalaryFormulaLIKE(TBL, Criteria + "%", CInt(vTipe))
                End If
            Else
                If SearchCriteria = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                    QueryResult = DTA.FillBy_DocumentNoLIKE2(TBL, Criteria + "%", CInt(vTipe), vUserCode)
                ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDate Then
                    QueryResult = DTA.FillBy_TanggalTransaksi2(TBL, vEndDate, vStartDate, CInt(vTipe), vUserCode)
                ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                    QueryResult = DTA.FillBy_RemarksLIKE2(TBL, Criteria + "%", CInt(vTipe), vUserCode)
                ElseIf SearchCriteria = EnumSearchCriteria.ByComponentName Then
                    QueryResult = DTA.FillBy_ComponentNameLIKE2(TBL, Criteria + "%", CInt(vTipe), vUserCode)
                ElseIf SearchCriteria = EnumSearchCriteria.BySalaryFormula Then
                    QueryResult = DTA.FillBy_SalaryFormulaLIKE2(TBL, Criteria + "%", CInt(vTipe), vUserCode)
                End If
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                'DGV.Columns(1).HeaderText = "Code"
                'DGV.Columns(2).HeaderText = "Description"
                'DGV.Columns(3).HeaderText = "City"
                'DGV.Columns(4).HeaderText = "Remarks"
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > TimeSpan.TicksPerSecond Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, _
                    0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal ij As FrmTrans_Gaji_Override) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_Payment_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, vObjCD, vPrimary, vdgvItem, ij)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal ij As FrmTrans_Gaji_Override) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_Payment_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, vPrimary, vObjCD, vdgvItem, ij)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vUserCode As String, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal ij As FrmTrans_Gaji_Override, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, vPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Loan_Payment_Hapus, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > ObjAntri.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If CInt(Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimary, vUserCode, True, TampilkanPesan, ij)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Sub GetNoAntrian(ByRef vIdentifier As String, ByVal ObjRnd As Random, _
    ByVal DTAInOUT As Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter, ByVal vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian, _
    ByRef vPrimary_Antrian As Double, ByVal vPrimaryTransaksiAsal As Double, ByVal vTipeTransaksiAsal As Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian, _
    ByVal vNamaUser As String)
        Dim NamaKomputer As String = ""
        Dim IP_Add As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjAntri As New Core_Module.ClsAntrian

        Try
            NamaKomputer = My.Computer.Name
            IP_Add = ObjCD.GetIpAddress_Local
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)

            vIdentifier = ObjRnd.Next(1000000)
            'jangan sampai sama dengan identifier yang sudah ada
            While DTAInOUT.ScalarQuery_GetJmlIdentifier(vTipeAntrian, vIdentifier).Value > 0
                vIdentifier = ObjRnd.Next(1000000)
            End While
            'masukan data baru ke antrian
            If System.Net.Dns.GetHostAddresses("").Length > 2 Then
                DTAInOUT.Insert(vTipeAntrian, vIdentifier, vPrimaryTransaksiAsal, vTipeTransaksiAsal, vNamaUser, My.Computer.Name + "-" + _
                System.Net.Dns.GetHostAddresses("")(2).ToString)
            Else
                DTAInOUT.Insert(vTipeAntrian, vIdentifier, vPrimaryTransaksiAsal, vTipeTransaksiAsal, vNamaUser, My.Computer.Name + "- localhost")
            End If

            'cek jangan sampai ada 2 data yang sama
            While DTAInOUT.ScalarQuery_GetJmlIdentifier(vTipeAntrian, vIdentifier).Value > 1
                'bila ada yang sama maka harus masukkan data baru
                DTAInOUT.DeleteQuery(vTipeAntrian, vIdentifier)
                'minta identifier baru
                vIdentifier = ObjRnd.Next(1000000)
                'jangan sampai sama dengan identifier yang sudah ada
                While DTAInOUT.ScalarQuery_GetJmlIdentifier(vTipeAntrian, vIdentifier).Value > 0
                    vIdentifier = ObjRnd.Next(1000000)
                    'masukan data baru ke antrian
                    DTAInOUT.Insert(vTipeAntrian, vIdentifier, vPrimaryTransaksiAsal, vTipeTransaksiAsal, vNamaUser, My.Computer.Name)
                End While
            End While

            'ambil no antrian
            vPrimary_Antrian = DTAInOUT.Scalar_Query_GetPrimary(vIdentifier, vTipeAntrian).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Function ApakahSudahAda(ByVal vDocumentNo As String) As Boolean
        Try
            If Me.GetJmlByDocumentNo(vDocumentNo) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetMAXUrutan(ByVal vEndDate As Date, ByVal vStartDate As Date, ByVal vTipe As EnumType) As Double
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetMAXUrutan(vEndDate, vStartDate, vTipe)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vObjCD As Core_Module.ClsCekData, ByVal vTipe As EnumType) As String
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim InOut As String = "OVR"
        NoUrut = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = vObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = vObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetMAXUrutan(TanggalEnd, TanggalStart, vTipe).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetMAXUrutan(TanggalEnd, TanggalStart, vTipe).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If


            'pastikan no urut terdiri atas 6 huruf
            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While

            If vTipe = EnumType.SalaryOverride Then
                InOut = "OVR"
            Else
                InOut = "OTH"
            End If
            Return NoUrut + "/" + InOut + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vPrimary As Double, _
    ByVal vShowMsg As Boolean, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_Gaji_Override) As Boolean
        Dim InfoString As String = ""
        Try
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mDocumentNo) = True Then
                    Me.mDocumentNo = Me.GetNoBaru(Me.mTanggalTransaksi, vObjCD, i.Tipe)
                End If
            Else
                Dim NoTransaksiLama As String = Me.GetDocumentNo(Me.mPrimary)
                If NoTransaksiLama <> Me.mDocumentNo Then
                    If Me.ApakahSudahAda(Me.mDocumentNo) = True Then
                        Me.mDocumentNo = Me.GetNoBaru(Me.mTanggalTransaksi, vObjCD, i.Tipe)
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. Minimal ada satu item 
            If i.dgvItem.Rows.Count <= 1 Then
                If vShowMsg = True Then
                    MsgBox("Please input data item.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '2. kode cash account
            If Me.mPrimaryFormula = "-1" Then
                If vShowMsg = True Then
                    MsgBox("Please input component name.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '3. Remarks
            If vObjCD.CekDataKosong(Me.mRemarks, False) = False Then
                If vShowMsg = True Then
                    MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vUserCode As String, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_Gaji_Override) As Boolean
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim ObjItem As New ClsTransaksiSalary_Override_Item

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, vPrimary, True, vObjCD, vDGVItem, i) = True Then
                Me.mUrutan = CDbl(Me.mDocumentNo.Substring(0, 6))
                If DTA.Insert(Me.mDocumentNo, Me.mTanggalTransaksi, Me.mFormulaType, _
                Me.mPrimaryFormula, Me.mIdentifierFormula, Me.mRemarks, _
                vUserCode, Now, vUserCode, Now, Now, Me.mUrutan, CInt(Me.mTipe)) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mDocumentNo)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVItem, i)

                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUserCode As String, _
    ByVal vPrimary As Double, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_Gaji_Override) As Boolean
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim ObjItem As New ClsTransaksiSalary_Override_Item
        Try
            If Me.ApakahDataBenar(False, vPrimary, True, vObjCD, vDGVItem, i) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                Me.mUrutan = CDbl(Me.mDocumentNo.Substring(0, 6))
                If DTA.Update(Me.mDocumentNo, Me.mTanggalTransaksi, _
                              Me.mFormulaType, Me.mPrimaryFormula, Me.mIdentifierFormula, Me.mRemarks, _
                              vUserCode, Now, Now, Me.mUrutan, CInt(Me.mTipe), Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVItem, i)

                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, ByVal i As FrmTrans_Gaji_Override) As Boolean
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim ObjItem As New ClsTransaksiSalary_Override_Item
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, vStrictDelete, vShowMsg, i) = True Then
                'Hapus item
                ObjItem.HapusItem(vPrimary)
                'Hapus
                DTA.Delete(vPrimary)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mDocumentNo = ""
            mUrutan = 0
            mTanggalTransaksi = Now
            mFormulaType = EnumFormulaType.Salary_Formula
            mPrimaryFormula = -1
            mIdentifierFormula = "-1-1"
            mRemarks = ""

            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today

            'Item
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictEdit As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim pCreatedBy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictEdit = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If

                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        MsgBox("User can edit this data: " + pCreatedBy, _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, ByVal i As FrmTrans_Gaji_Override) As Boolean
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim pCreatedBy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictDelete = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If
                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        MsgBox("User that can delete this data is: " + pCreatedBy, _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub Cetakan(ByVal vPrimary As Double, ByVal vDef_Prec As Double, _
    ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vTipe As EnumType)
        Dim DS As New DseTransaksiSalary_Override
        Dim Report As New CrpTransSalary_Override_Amount
        Dim DTA As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount_item_1TableAdapter
        Dim DTAHeader As New DseTransaksiSalary_OverrideTableAdapters.tbltrans_salary_override_amount1TableAdapter
        Dim TBL As DseTransaksiSalary_Override.tbltrans_salary_override_amount_item_1DataTable = DS.tbltrans_salary_override_amount_item_1
        Dim TBLHeader As New DseTransaksiSalary_Override.tbltrans_salary_override_amount1DataTable
        Dim RowHeader As DseTransaksiSalary_Override.tbltrans_salary_override_amount1Row
        Dim ObjFormula As New Core_Module.ClsMasterSalaryFormula
        Dim ObjItemFormula As New Core_Module.ClsMasterSalaryFormula_item
        Dim ObjOther As New Core_Module.ClsTrans_OtherDeduction
        Dim mParamDocNo As String = ""
        Dim mParamDate As String = ""
        Dim mParamFormulaType As String = ""
        Dim mParamComponentName As String = ""
        Dim mParamSalaryFormulaCode As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAHeader.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryTransaksi(TBL, vPrimary)
            DTAHeader.FillBy_Primary(TBLHeader, vPrimary)

            If TBLHeader.Rows.Count > 0 Then
                RowHeader = TBLHeader.Rows(0)
                mParamDocNo = RowHeader.Document_No
                mParamDate = Format(RowHeader._Date, "dd MMM yyyy")
                If RowHeader.Formula_Type = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula Then
                    mParamFormulaType = "Salary formula"
                    mParamComponentName = ObjItemFormula.GetDescComponent(RowHeader.Identifier_Formula)
                    mParamSalaryFormulaCode = ObjFormula.GetCodeByPrimary(RowHeader.Primary_Formula)
                ElseIf RowHeader.Formula_Type = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula_Other_Deduction Then
                    mParamFormulaType = "Other deduction"
                    mParamComponentName = ObjOther.GetCodeByPrimary(RowHeader.Primary_Formula)
                    mParamSalaryFormulaCode = ObjOther.GetCodeByPrimary(RowHeader.Primary_Formula)
                End If
            End If

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, vDef_Prec)
            Report.SetParameterValue(Report.Parameter_DocumentNo.ParameterFieldName, mParamDocNo)
            Report.SetParameterValue(Report.Parameter_Date.ParameterFieldName, mParamDate)
            Report.SetParameterValue(Report.Parameter_FormulaType.ParameterFieldName, mParamFormulaType)
            Report.SetParameterValue(Report.Parameter_Component_Name.ParameterFieldName, mParamComponentName)
            Report.SetParameterValue(Report.Parameter_FormulaCode.ParameterFieldName, mParamSalaryFormulaCode)
            If vTipe = EnumType.SalaryOverride Then
                Report.SetParameterValue(Report.Parameter_Judul.ParameterFieldName, "Override Salary")
            Else
                Report.SetParameterValue(Report.Parameter_Judul.ParameterFieldName, "Other Payment")
            End If

            vCRV.ReportSource = Report
            vCRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAHeader.Dispose()
            TBLHeader.Dispose()
            ObjFormula = Nothing
            ObjItemFormula = Nothing
            ObjOther = Nothing
        End Try
    End Sub

End Class
