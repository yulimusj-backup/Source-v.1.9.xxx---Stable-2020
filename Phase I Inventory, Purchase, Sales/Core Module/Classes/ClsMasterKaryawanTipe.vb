Option Explicit On
Option Strict Off
Public Class ClsMasterKaryawanTipe
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mEmployeetypeCode As String = ""
    Dim mDescription As String = ""
    Dim mRemarks As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property EmployeetypeCode() As String
        Get
            Return Me.mEmployeetypeCode
        End Get
        Set(ByVal value As String)
            mEmployeetypeCode = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return Me.mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
        End Set
    End Property
    Public Property Remarks() As String
        Get
            Return Me.mRemarks
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
    Public Function ApakahSudahAda(ByVal vEmployeetypeCode As String) As Boolean
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_tipekontrakTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vEmployeetypeCode).HasValue Then
                If DTA.ScalarQuery_GetJml(vEmployeetypeCode).Value > 0 Then
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
    Public Function GetTypeCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_tipekontrakTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_getTypeCode(vPrimary)) = False Then
                Return DTA.ScalarQuery_getTypeCode(vPrimary)
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
    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_tipekontrakTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_getTimestamp(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetPrimary(ByVal vCode As String) As String
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_tipekontrakTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_getPrimary(vCode)) = False Then
                Return DTA.ScalarQuery_getPrimary(vCode)
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
    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_tipekontrakTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_getTypeCode(vPrimary)) = False Then
                Return DTA.ScalarQuery_getTypeCode(vPrimary)
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
    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_tipekontrakTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_getDesc(vPrimary)) = False Then
                Return DTA.ScalarQuery_getDesc(vPrimary)
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
    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_tipekontrakTableAdapter
        Dim TBL As New DseMasterKaryawanTipe.tblmaster_karyawan_tipekontrakDataTable
        Dim row As DseMasterKaryawanTipe.tblmaster_karyawan_tipekontrakRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = row.Primary
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mDescription = .Description
                    Me.mRemarks = .Remarks
                    Me.mModifiedBy = .Modifed_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mEmployeetypeCode = .Code
                    Me.mTimestamp = .Timestamp
                End With
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function
    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_tipekontrakTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mEmployeetypeCode, Me.mDescription, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_getPrimary(Me.mEmployeetypeCode).Value
                Me.mTimestamp = DTA.ScalarQuery_getTimestamp(Me.mPrimary).Value
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode
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
    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblCarimaster_karyawan_tipekontrakTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_getJmlByCode(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_getJmlByCode(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_getJmlByDesc(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_getJmlByDesc(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                If DTA.ScalarQuery_getJmlByRemarks(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_getJmlByRemarks(vKriteria + "%").Value
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
    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_tipekontrakTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mEmployeetypeCode, Me.mDescription, Me.Remarks, vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_getTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode
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

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterKaryawanTipe.tblCarimaster_karyawan_tipekontrakDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblCarimaster_karyawan_tipekontrakTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescriptionLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_remarksLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Contract Type"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
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
    Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean
        Dim Obj As New ClsMasterKaryawanTipe

        Try
            'tidak boleh dihapus bila sudah ada yang merefer:
            '1. customer
            If Obj.GetJmlBy_EmployeeType(vPrimary) > 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Function
    Public Function ApakahBolehEdit(ByVal vPrimary As Double) As Boolean
        Dim Obj As New ClsMasterKaryawanTipe

        Try
            'tidak boleh diEdit bila sudah ada yang merefer:
            '1. customer
            If Obj.GetJmlBy_EmployeeType(vPrimary) > 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Function
    Public Function GetJmlBy_EmployeeType(ByVal vPrimaryEmployeeType As Double) As Integer
        Dim dta As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_item_kontrakTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_getJml(vPrimaryEmployeeType).HasValue Then
                Return dta.ScalarQuery_getJml(vPrimaryEmployeeType).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function
    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_tipekontrakTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.Delete(vPrimary)
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
    Public Function isiCboEmployeeTypeContrack(ByVal TBL As DseMasterKaryawanTipe.tblmaster_karyawan_tipekontrakDataTable) As Boolean
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_tipekontrakTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(TBL) > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Public Function GetIndexItemKontrakTipe(ByVal vDocument As String, _
           ByVal vRowIdxFocus As Integer, ByVal vFrm As FrmMasterKaryawan) As Double
        Dim RowDGV As DataGridViewRow
        Try
            'pastikan tidak ada code yang sama
            For ij As Integer = 0 To vFrm.DGVKontrak.Rows.Count - 1
                RowDGV = vFrm.DGVKontrak.Rows(ij)
                If RowDGV.Cells(vFrm.ColContract_DocNo.Index).Value <> "" Then
                    If vRowIdxFocus <> ij Then
                        If RowDGV.Cells(vFrm.ColContract_DocNo.Index).Value = vDocument Then
                            Return vFrm.DGVKontrak.Rows.IndexOf(RowDGV)
                        End If
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Public Function CekValueDate(ByVal vFrm As FrmMasterKaryawan) As Boolean

        Try
            'cek baris atasnya
            If vFrm.DGVKontrak.CurrentRow.Index <> 0 Then
                If vFrm.DGVKontrak.Item(vFrm.ColContract_Start.Index, vFrm.DGVKontrak.CurrentRow.Index).Value <= vFrm.DGVKontrak.Item(vFrm.ColContract_End.Index, vFrm.DGVKontrak.CurrentRow.Index - 1).Value Then
                    MsgBox("The Start Date value for this row must be greater than the End date value in the previous row.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            'cek baris bawahnya
            If vFrm.DGVKontrak.Item(vFrm.ColContract_Start.Index, vFrm.DGVKontrak.CurrentRow.Index + 1).Value <> Nothing Then
                If vFrm.DGVKontrak.Item(vFrm.ColContract_End.Index, vFrm.DGVKontrak.CurrentRow.Index).Value >= vFrm.DGVKontrak.Item(vFrm.ColContract_Start.Index, vFrm.DGVKontrak.CurrentRow.Index + 1).Value Then
                    MsgBox("The End date value for this row must be smaller than the Start date value in the next row.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Public Function GetIdentifierBaru(ByVal vFrm As FrmMasterKaryawan) As String
        Dim IdentifierD As Integer = 0
        Dim NilaiMax As Integer = 0
        Dim Retval As String = ""

        Try
            For Each rowDGV As DataGridViewRow In vFrm.DGVKontrak.Rows
                IdentifierD = rowDGV.Cells(vFrm.ColContract_IdentifierContract.Index).Value.ToString.IndexOf("-", 1)
                IdentifierD = rowDGV.Cells(vFrm.ColContract_IdentifierContract.Index).Value.ToString.Substring(IdentifierD + 1)
                If NilaiMax < IdentifierD Then
                    NilaiMax = IdentifierD
                End If
            Next
            NilaiMax += 1

            Retval = "-1" & "-" & NilaiMax.ToString
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
    Public Sub TambahBarisDGVKontrak(ByVal vTblkontrak As DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrakDataTable, _
    ByVal vPrimaryTypeContract As Double, ByVal vPrimaryKaryawan As String, _
    ByVal vDocNo As String, ByVal vDate As String, ByVal vStart As String, ByVal vEnd As String, _
     ByVal vRemarks As String, ByVal vIdentifier As String, ByVal vFrm As FrmMasterKaryawan)
        Dim rowItemTipeKontrak As DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrakRow
        Try
            rowItemTipeKontrak = vFrm.tblItemTipeKontrak.Newtblmaster_karyawan_item_kontrakRow
            With rowItemTipeKontrak
                .primary = -1
                .primary_tipe_kontrak = vPrimaryTypeContract
                .primary_karyawan = vPrimaryKaryawan
                .document_no = vDocNo
                If vDate = "" Then
                    ._date = Format(Now, "dd MMM yyyy")
                Else
                    ._date = vDate
                End If
                If vStart = "" Then
                    .start_date = Format(Now, "dd MMM yyyy")
                Else
                    .start_date = vStart
                End If
                If vEnd = "" Then
                    .end_date = Format(Now, "dd MMM yyyy")
                Else
                    .end_date = vEnd
                End If
                .remarks = vRemarks
                .identifier_kontrak = vIdentifier
            End With
            vTblkontrak.Rows.Add(rowItemTipeKontrak)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Sub TambahBarisDGVAttchKontrak(ByVal vTblAttchkontrak As DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrak_attachmentDataTable, _
    ByVal vPrimaryKaryawan As String, ByVal vDesc As String, ByVal vFilePath As String, ByVal vRemarks As String, ByVal vIdentifier As String, _
    ByVal vFrm As FrmMasterKaryawan)
        Dim rowItemTipeKontrakAttch As DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrak_attachmentRow
        Try
            rowItemTipeKontrakAttch = vFrm.tblItemAttch.Newtblmaster_karyawan_item_kontrak_attachmentRow
            With rowItemTipeKontrakAttch
                .primary = -1
                .primary_karyawan = vPrimaryKaryawan
                .desc = vDesc
                If IsNothing(vFilePath) = True Then
                    .file_path = ""
                Else

                    .file_path = vFilePath
                End If
                If IsNothing(vRemarks) = True Then
                    .remarks = ""
                Else

                    .remarks = vRemarks
                End If

                .identifier_kontrak = vIdentifier
            End With
            vTblAttchkontrak.Rows.Add(rowItemTipeKontrakAttch)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
