Option Explicit On
Option Strict Off

Public Class ClsMasterTaxMeasurementReference
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mCode As String = ""
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

    Public Property Code() As String
        Get
            Return Me.mCode
        End Get
        Set(ByVal value As String)
            mCode = value
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

    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mCode, Me.mDescription, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCode).Value
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
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

    Public Function GetJml(ByVal vCode As String) As Double
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vCode).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vCode).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblcariTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                If DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").Value
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
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Update(Me.mCode, Me.mDescription, Me.mRemarks, vUserCode, Now, Now, CDec(Me.mPrimary), Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
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

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.Delete(CDec(vPrimary))
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

    Public Function ApakahSudahAda(ByVal vmCode As String) As Boolean
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter

        Try
            If DTA.ScalarQuery_GetJml(vmCode).HasValue Then
                If DTA.ScalarQuery_GetJml(vmCode).Value > 0 Then
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

    Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean
        Try
            'tidak boleh dihapus bila sudah ada yang merefer:
            '1. Master unit
            If Me.GetJmlMasterUnitByTaxUnitReference(vPrimary) > 0 Then
                MsgBox("You cannot delete this data, because there are master unit using this data.", _
                                 MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterTaxMeasurementReference.tblcariDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblcariTableAdapter
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
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Code"
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

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter
        Dim TBL As New DseMasterTaxMeasurementReference.tblmaster_taxmeasurement_referenceDataTable
        Dim row As DseMasterTaxMeasurementReference.tblmaster_taxmeasurement_referenceRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(Primary))

            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterTaxMeasurementReference.tblmaster_taxmeasurement_referenceRow)
                With row
                    Me.mPrimary = row.primary
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mDescription = .Description
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mCode = .code
                    Me.mRemarks = .Remarks
                    Me.mTimestamp = .timestamp
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

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCode(CDec(vPrimary))) = False Then
                Return DTA.ScalarQuery_GetCode(CDec(vPrimary))
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
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc(CDec(vPrimary))) = False Then
                Return DTA.ScalarQuery_GetDesc(CDec(vPrimary))
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
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(CDec(vPrimary)).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vmCode As String) As Double
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vmCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vmCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlMasterUnitByTaxUnitReference(ByVal vPrimaryTaxUnit As Double) As Double
        Dim DTA As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter
        Dim r As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            r = DTA.ScalarQuery_GetJmlUnitByTaxUnit(vPrimaryTaxUnit)

            If r.HasValue = True Then
                If r.Value > 0 Then
                    Return r.Value
                End If
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetInsertString() As String
        Dim retval As String = ""
        Try
            retval = "INSERT INTO `tblmaster_taxmeasurement_reference` (`Code`,`Description`,`Remarks`) VALUES" & _
            "('UM.0003','Kilogram',''),('UM.0004','Gram',''),('UM.0005','Karat',''),('UM.0001','Metrik Ton',''),('UM.0002','Wet Ton',''),('UM.0006','Kiloliter','')," & _
            "('UM.0007','Liter',''),('UM.0008','Barrel',''),('UM.0009','MMBTU',''),('UM.0010','Ampere',''),('UM.0011','Sentimeter Kubik',''),('UM.0012','Meter Persegi',''),('UM.0013','Meter','')," & _
            "('UM.0014','Inches',''),('UM.0015','Sentimeter',''),('UM.0016','Yard',''),('UM.0017','Lusin',''),('UM.0018','Unit',''),('UM.0019','Set',''),('UM.0020','Lembar','')," & _
            "('UM.0021','Piece',''),('UM.0022','Boks',''),('UM.0023','Tahun',''),('UM.0024','Bulan',''),('UM.0025','Minggu',''),('UM.0026','Hari',''),('UM.0027','Jam',''), " & _
            "('UM.0028','Menit',''),('UM.0029','Persen',''),('UM.0030','Kegiatan',''),('UM.0031','Laporan',''),('UM.0032','Bahan',''),('UM.0033','Lainnya','');"

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function SimpanItem_UsingScript(ByVal i As FrmMasterTaxMeasurementReference) As Boolean
        Dim conStr As String = My.Settings.Core_DbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Try
            'open connection
            Conn.Open()
            If Me.ApakahSudahAdaData = False Then
                sql = Me.GetInsertString()
                script.Query = sql
                'execute script
                script.Execute()
                MsgBox("Insert data completed.", MsgBoxStyle.Information, "Information")
            End If

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function ApakahSudahAdaData() As Boolean
        Dim d As New DseMasterTaxMeasurementReferenceTableAdapters.tblmaster_taxmeasurement_referenceTableAdapter
        Dim l As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            l = d.ScalarQuery_GetJmlTotal()
            If l.HasValue = True Then
                If l.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            d.Dispose()
        End Try
    End Function
End Class
