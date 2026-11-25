Option Explicit On
Option Strict Off

Public Class ClsMasterJobPosition
    Private mPrimary As Double = -1
    Private mCode As String = ""
    Private mDesc As String = ""
    Private mRemarks As String = ""
    Private mJobDesc As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim vTipeTransaksiAsal As Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian = ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterProses

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property JobDesc() As String
        Get
            Return Me.mJobDesc
        End Get
        Set(ByVal value As String)
            mJobDesc = value
        End Set
    End Property

    Public Property Code() As String
        Get
            Return mCode
        End Get
        Set(ByVal value As String)
            mCode = value
        End Set
    End Property

    Public Property Desc() As String
        Get
            Return mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
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
            Return Me.mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return Me.mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return Me.mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return Me.mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return Me.mTimestamp
        End Get
        Set(ByVal value As Date)
            Me.mTimestamp = value
        End Set
    End Property

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
        Dim Dummy As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dummy = DTA.ScalarQuery_GetPrimary(vCode)
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

    Public Function GetPrimary_ByDesc(ByVal vDesc As String) As Double
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
        Dim Dummy As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dummy = DTA.ScalarQuery_GetPrimary(vDesc)
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

    Public Function GetPeriodic(ByVal vPrimaryJob As Double, ByVal vPrimaryKompetensi As Double) As String
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobposition_item_kompetensi1TableAdapter
        Dim dummy As Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dummy = DTA.ScalarQuery_GetPeriodicUpdateItem(vPrimaryJob, vPrimaryKompetensi)
            If dummy.HasValue Then
                Return dummy
            Else
                Return "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetValue_CustomFormatted(ByVal vPrimaryJob As Double, ByVal vPrimaryKompetensi As Double, ByVal vObjCD As Core_Module.ClsCekData, _
  ByVal vDataTipe As ClsMasterKompetensi.EnumDataType) As String
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobposition_item_kompetensi1TableAdapter
        Dim Def_Prec As Integer = 0
        Dim dummy As String
        Dim ObjKompetensi As New ClsMasterKompetensi
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dummy = DTA.ScalarQuery_GetValueByJobNkompetensi(vPrimaryJob, vPrimaryKompetensi)
            If String.IsNullOrEmpty(dummy) = False Then
                Def_Prec = ObjKompetensi.GetDefPrec(vPrimaryKompetensi)
                Return ObjKompetensi.SetFormatDataCustom(dummy, vDataTipe, Def_Prec, vObjCD)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjKompetensi = Nothing
        End Try
    End Function

    Public Function GetJmlItem(ByVal vPrimaryJob As Double, ByVal vPrimaryKompetensi As Double) As Double
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobposition_item_kompetensi1TableAdapter
        Dim dummy As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dummy = DTA.ScalarQuery_GetJmlKompetensiByJob(vPrimaryJob, vPrimaryKompetensi)
            If dummy.HasValue = True Then
                Return dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetCode(vPrimary)

            If String.IsNullOrEmpty(Retval) = False Then
                Return Retval
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

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
        Dim Dummy As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTimestamp(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
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

    Public Function ApakahMemilikiKompetensi(ByVal vPrimary As String) As Boolean
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobposition_item_kompetensiTableAdapter
        Dim Dummy As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJml(vPrimary)
            If Dummy.HasValue Then
                If Dummy.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
        Dim Dummy As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByCode(vCode)
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

    ''' <summary>
    ''' data ini digunakan untuk mengecek kebenaran data yang diberikan
    ''' </summary>
    ''' <param name="vBaru"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vShowMsg As Boolean, ByVal vDGV As DataGridView, _
    ByVal vFrm As FrmMasterJobPosition) As Boolean
        Dim ObjKompetensi As New ClsMasterKompetensi

        Try
            'tidak boleh ada kode yang sama
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mCode) = True Then
                    If vShowMsg = True Then
                        MsgBox("Code duplication, please change code.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.GetCode(Me.mPrimary)
                If CodeLama <> Me.mCode Then
                    If Me.ApakahSudahAda(Me.mCode) = True Then
                        If vShowMsg = True Then
                            MsgBox("Code duplication, please change code.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. kode
            If Me.mCode = "" Then
                If vShowMsg = True Then
                    MsgBox("Please specify code.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '2. nama
            If Me.mDesc = "" Then
                If vShowMsg = True Then
                    MsgBox("Please specify description.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '3. pastikan semua itemnya memiliki pasangan tipe data dan primary kompetensi yang tepat
            If vDGV.RowCount > 1 Then
                If ObjKompetensi.ApakahTipeDataTepat(vDGV, vFrm.ColPrimaryCompetence.Index, vFrm.ColValueType.Index, _
                            vFrm.ColCode.Index, vShowMsg) = False Then
                    Return False
                End If
            End If


            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjKompetensi = Nothing
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vStrictEdit As Boolean, _
    ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
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

                'If pCreatedBy <> vUserCode Then
                'If vShowMsg = True Then
                'MsgBox("User yang bisa mengedit transaksi ini adalah: " + pCreatedBy, MsgBoxStyle.Information, "Information")
                'End If
                'Return False
                'End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vStrictDelete As Boolean, _
    ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
        Dim pCreatedBy As String = ""
        Dim DTAKaryawan As New DseMasterJobPositionTableAdapters.tblmaster_karyawanTableAdapter
        Dim dummy As Nullable(Of Long)

        Dim TipeParameter As ClsMasterSalaryFormula_item.EnumTipeConstraitVariable = ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Kompetensi
        Dim ObjSal As New ClsMasterSalaryFormula

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAKaryawan.Connection.ConnectionString = DTA.Connection.ConnectionString

            '1.tidak boleh dihapus bila sudah ada yang merefer:
            'a. karyawan
            dummy = DTAKaryawan.ScalarQuery_GetJmlByPrimaryJobPosition(vPrimary)
            If dummy.HasValue = True Then
                If dummy.Value > 0 Then
                    MsgBox("This data used on employee data, you cannot delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If ObjSal.GetJmlConstraint(131, 3, ClsMasterSalaryFormula_item.EnumTipeConstraitVariable.Hardcode, vPrimary.ToString) > 0 Then
                MsgBox("There salary formula using this data, You can't delete this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAKaryawan.Dispose()
        End Try
    End Function

    ''
    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterJobPosition.tblmaster_jobposition1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobposition1TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCompetenceCode Then
                QueryResult = DTA.FillBy_CompetenceCode(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCompetenceDesc Then
                QueryResult = DTA.FillBy_CompetenceDesc(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
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

    Public Function CariDetail(ByVal Primary As Double, ByVal vDGVItem As DataGridView, ByVal vFrm As FrmMasterJobPosition) As Boolean
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
        Dim TBL As New DseMasterJobPosition.tblmaster_jobpositionDataTable
        Dim row As DseMasterJobPosition.tblmaster_jobpositionRow
        Dim DTAItem As New DseMasterJobPositionTableAdapters.tblmaster_jobposition_item_kompetensi1TableAdapter
        Dim TBLItem As New DseMasterJobPosition.tblmaster_jobposition_item_kompetensi1DataTable
        Dim rowItem As DseMasterJobPosition.tblmaster_jobposition_item_kompetensi1Row
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New ClsCekData
        Dim DummyBool As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.FillBy_Primary(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mCode = .code
                    Me.mDesc = .desc
                    Me.mRemarks = .remarks
                    Me.mJobDesc = .job_descrption

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
            End If

            'isikan data item
            DummyBool = vDGVItem.AllowUserToAddRows
            vDGVItem.AllowUserToAddRows = False
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.Rows.Clear()

            DTAItem.Fill(TBLItem, Me.mPrimary)
            For Each rowItem In TBLItem.Rows
                vDGVItem.Rows.Add()
                rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)

                With rowDGV
                    .Cells(vFrm.ColCode.Index).Value = rowItem.Competence_Code
                    .Cells(vFrm.ColDesc.Index).Value = rowItem.Description
                    .Cells(vFrm.ColNo.Index).Value = vDGVItem.Rows.Count
                    .Cells(vFrm.ColPrimaryCompetence.Index).Value = rowItem.primary_kompetensi
                    If rowItem.periodic_update = 0 Then
                        .Cells(vFrm.ColPeriodicUpdate.Index).Value = Boolean.FalseString
                    Else
                        .Cells(vFrm.ColPeriodicUpdate.Index).Value = Boolean.TrueString
                    End If
                    .Cells(vFrm.ColRemarks.Index).Value = rowItem.remarks
                    .Cells(vFrm.ColValueType.Index).Value = rowItem.Data_Type

                    If rowItem.Data_Type = ClsMasterKompetensi.EnumDataType.Number Then
                        .Cells(vFrm.ColValue.Index).Value = ObjCD.DigitGrouping(rowItem.value, rowItem.Default_Precision)
                        .Cells(vFrm.ColDefaultValue.Index).Value = ObjCD.DigitGrouping(rowItem.Default_Value, rowItem.Default_Precision)
                    ElseIf rowItem.Data_Type = ClsMasterKompetensi.EnumDataType.Text Then
                        .Cells(vFrm.ColDefaultValue.Index).Value = rowItem.Default_Value
                        .Cells(vFrm.ColValue.Index).Value = rowItem.value
                    ElseIf rowItem.Data_Type = ClsMasterKompetensi.EnumDataType.YN Then
                        If rowItem.Default_Value = 0 Then
                            .Cells(vFrm.ColDefaultValue.Index).Value = "No"
                        Else
                            .Cells(vFrm.ColDefaultValue.Index).Value = "Yes"
                        End If
                        If rowItem.value = 0 Then
                            .Cells(vFrm.ColValue.Index).Value = "No"
                        Else
                            .Cells(vFrm.ColValue.Index).Value = "Yes"
                        End If
                    ElseIf rowItem.Data_Type = ClsMasterKompetensi.EnumDataType.Tanggal Then
                        .Cells(vFrm.ColDefaultValue.Index).Value = Format(CDate(rowItem.Default_Value), "dd MMM yyyy")
                        .Cells(vFrm.ColValue.Index).Value = Format(CDate(rowItem.value), "dd MMM yyyy")
                    ElseIf rowItem.Data_Type = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                        If rowItem.value = ClsMasterKompetensi.Aphabetic.A Then
                            .Cells(vFrm.ColValue.Index).Value = "A"
                        ElseIf rowItem.value = ClsMasterKompetensi.Aphabetic.B Then
                            .Cells(vFrm.ColValue.Index).Value = "B"
                        ElseIf rowItem.value = ClsMasterKompetensi.Aphabetic.C Then
                            .Cells(vFrm.ColValue.Index).Value = "C"
                        ElseIf rowItem.value = ClsMasterKompetensi.Aphabetic.D Then
                            .Cells(vFrm.ColValue.Index).Value = "D"
                        ElseIf rowItem.value = ClsMasterKompetensi.Aphabetic.E Then
                            .Cells(vFrm.ColValue.Index).Value = "E"
                        ElseIf rowItem.value = ClsMasterKompetensi.Aphabetic.F Then
                            .Cells(vFrm.ColValue.Index).Value = "F"
                        End If

                        If rowItem.Default_Value = ClsMasterKompetensi.Aphabetic.A Then
                            .Cells(vFrm.ColDefaultValue.Index).Value = "A"
                        ElseIf rowItem.Default_Value = ClsMasterKompetensi.Aphabetic.B Then
                            .Cells(vFrm.ColDefaultValue.Index).Value = "B"
                        ElseIf rowItem.Default_Value = ClsMasterKompetensi.Aphabetic.C Then
                            .Cells(vFrm.ColDefaultValue.Index).Value = "C"
                        ElseIf rowItem.Default_Value = ClsMasterKompetensi.Aphabetic.D Then
                            .Cells(vFrm.ColDefaultValue.Index).Value = "D"
                        ElseIf rowItem.Default_Value = ClsMasterKompetensi.Aphabetic.E Then
                            .Cells(vFrm.ColDefaultValue.Index).Value = "E"
                        ElseIf rowItem.Default_Value = ClsMasterKompetensi.Aphabetic.F Then
                            .Cells(vFrm.ColDefaultValue.Index).Value = "F"
                        End If
                    End If
                End With
            Next
            vDGVItem.AllowUserToAddRows = DummyBool
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vKode As String) As Double
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByCode(vKode).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByCode(vKode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As EnumSearchCriteria, ByVal vStatusAktif As Boolean, _
    ByVal vSearchAll As Boolean) As Double
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobposition1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchBy = EnumSearchCriteria.ByCode Then
                Return DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByCompetenceCode Then
                Return DTA.ScalarQuery_GetJmlByCompetenceCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByCompetenceDesc Then
                Return DTA.ScalarQuery_GetJmlByCompetenceDescLIKE(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub New()
        Try
            Me.DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vDGV As DataGridView, ByVal vFrm As FrmMasterJobPosition) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, -1, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterJobPosition, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                           retval = Me.SimpanBaru1(vUserCode, vDGV, vFrm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vDGV As DataGridView, ByVal vFrm As FrmMasterJobPosition) As Boolean
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(True, True, vDGV, vFrm) = True Then
                If DTA.Insert(Me.mCode, Me.mDesc, vUserCode, Now, vUserCode, Now, Now, Me.mJobDesc, Me.mRemarks) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCode)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan item
                    Me.SimpanItem(vDGV, vFrm)

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
        End Try
    End Function

    Private Sub SimpanItem(ByVal vDGV As DataGridView, ByVal vFrm As FrmMasterJobPosition)
        Dim rowDGV As DataGridViewRow
        Dim DummyBool As Boolean = False
        Dim DTAItem As New DseMasterJobPositionTableAdapters.tblmaster_jobposition_item_kompetensiTableAdapter
        Dim Value As String = ""
        Dim ValueType As ClsMasterKompetensi.EnumDataType = ClsMasterKompetensi.EnumDataType.Alfabetic
        Dim vPeriodicUpdate As String = ""

        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring

            DummyBool = vDGV.AllowUserToAddRows
            vDGV.AllowUserToAddRows = False
            For Each rowDGV In vDGV.Rows
                ValueType = rowDGV.Cells(vFrm.ColValueType.Index).Value
                If ValueType = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                    If rowDGV.Cells(vFrm.ColValue.Index).Value = "A" Then
                        Value = CInt(ClsMasterKompetensi.Aphabetic.A)
                    ElseIf rowDGV.Cells(vFrm.ColValue.Index).Value = "B" Then
                        Value = CInt(ClsMasterKompetensi.Aphabetic.B)
                    ElseIf rowDGV.Cells(vFrm.ColValue.Index).Value = "C" Then
                        Value = CInt(ClsMasterKompetensi.Aphabetic.C)
                    ElseIf rowDGV.Cells(vFrm.ColValue.Index).Value = "D" Then
                        Value = CInt(ClsMasterKompetensi.Aphabetic.D)
                    ElseIf rowDGV.Cells(vFrm.ColValue.Index).Value = "E" Then
                        Value = CInt(ClsMasterKompetensi.Aphabetic.E)
                    ElseIf rowDGV.Cells(vFrm.ColValue.Index).Value = "F" Then
                        Value = CInt(ClsMasterKompetensi.Aphabetic.F)
                    End If
                ElseIf ValueType = ClsMasterKompetensi.EnumDataType.Number Or _
                ValueType = ClsMasterKompetensi.EnumDataType.Text Then
                    Value = rowDGV.Cells(vFrm.ColValue.Index).Value
                ElseIf ValueType = ClsMasterKompetensi.EnumDataType.Tanggal Then
                    Value = Format(CDate(rowDGV.Cells(vFrm.ColValue.Index).Value), "MM/dd/yyyy")
                ElseIf ValueType = ClsMasterKompetensi.EnumDataType.YN Then
                    If rowDGV.Cells(vFrm.ColValue.Index).Value.ToString.ToLower = "yes" Then
                        Value = 1
                    Else
                        Value = 0
                    End If
                End If
                If CBool(rowDGV.Cells(vFrm.ColPeriodicUpdate.Index).Value) = False Then
                    vPeriodicUpdate = 0
                Else
                    vPeriodicUpdate = 1
                End If

                DTAItem.Insert(Me.mPrimary, rowDGV.Cells(vFrm.ColPrimaryCompetence.Index).Value, _
                Value, rowDGV.Cells(vFrm.ColRemarks.Index).Value, _
                vPeriodicUpdate)
            Next

            vDGV.AllowUserToAddRows = DummyBool
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
        End Try
    End Sub

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vDGV As DataGridView, ByVal vFrm As FrmMasterJobPosition) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterJobPosition, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, vDGV, vFrm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal vDGV As DataGridView, ByVal vFrm As FrmMasterJobPosition) As Boolean
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
        Dim DTAItem As New DseMasterJobPositionTableAdapters.tblmaster_jobposition_item_kompetensiTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            If Me.ApakahDataBenar(False, True, vDGV, vFrm) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTAItem.Connection.ConnectionString = DTA.Connection.ConnectionString

                If DTA.Update(Me.mCode, Me.mDesc, vUserCode, Now, Now, Me.mJobDesc, Me.mRemarks, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus item lama
                    DTAItem.Delete(Me.mPrimary)

                    'simpan item
                    Me.SimpanItem(vDGV, vFrm)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAItem.Dispose()
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
             Me.vPrimary_Antrian, vPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterJobPosition, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimary, vUserCode, _
                            vStrictDelete, vShowMsg)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahBolehHapus(vPrimary, vUserCode, vStrictDelete, vShowMsg) = True Then
                DTA.Delete(vPrimary) 'itemnya hilang berdasarkan cascading

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterJobPositionTableAdapters.tblmaster_jobpositionTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetDesc(vPrimary)
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
End Class
