Option Explicit On
Option Strict Off

Public Class ClsMasterSalesPerson
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mDescription As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim mPrimaryGroup As Double = -1

    Public Property PrimaryGroup() As Double
        Get
            Return Me.mPrimaryGroup
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryGroup = value
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

    Public Property Code() As String
        Get
            Return Me.mCode
        End Get
        Set(ByVal value As String)
            Me.mCode = value
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
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.mCode, Me.mDescription, Now, vUserCode, Now, vUserCode, Now, Me.mPrimaryGroup) > 0 Then
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
       EnumSearchCriteria) As Double
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                If DTA.ScalarQuery_GetJml1(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml1(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                If DTA.ScalarQuery_GetJml2(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml2(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByGroupCode Then
                If DTA.ScalarQuery_GetJml3(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml3(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByGroupDesc Then
                If DTA.ScalarQuery_GetJml4(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJml4(vKriteria + "%").Value
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
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Update(Me.mCode, Me.mDescription, Now, vUserCode, Now, Me.mPrimaryGroup, CDec(Me.mPrimary), Me.mTimestamp) > 0 Then
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
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter

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

    Public Function ApakahSudahAda(ByVal vSalesPersonCode As String) As Boolean
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vSalesPersonCode).HasValue Then
                If DTA.ScalarQuery_GetJml(vSalesPersonCode).Value > 0 Then
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
        Dim Obj As New clsMasterProduct
        Dim DTA_Cust As New DseMasterSalesPersonTableAdapters.tblmaster_customerTableAdapter
        Dim DTA_SO As New DseMasterSalesPersonTableAdapters.tbltrans_soTableAdapter
        Dim DTA_Inv As New DseMasterSalesPersonTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim DTA_Inv_Obs As New DseMasterSalesPersonTableAdapters.tbltrans_salesinvoice_obseleteTableAdapter
        Dim DTABudget As New DseMasterSalesPersonTableAdapters.tbltrans_salesbudget_itemTableAdapter

        Try
            DTA_Cust.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Inv.Connection.ConnectionString = DTA_Cust.Connection.ConnectionString
            DTA_Inv_Obs.Connection.ConnectionString = DTA_Cust.Connection.ConnectionString
            DTA_SO.Connection.ConnectionString = DTA_Cust.Connection.ConnectionString
            DTABudget.Connection.ConnectionString = DTA_Cust.Connection.ConnectionString

            'tidak bisa dihapus, jika sudah digunakan pada master customer
            If DTA_Cust.ScalarQuery_GetJml(vPrimary).HasValue = True Then
                If DTA_Cust.ScalarQuery_GetJml(vPrimary).Value > 0 Then
                    MsgBox("There customer refering to this data, you cannot delete this data", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'tidak bisa dihapus, jika sudah digunakan pada transaksi so
            If DTA_SO.ScalarQuery_GetJml(vPrimary).HasValue = True Then
                If DTA_SO.ScalarQuery_GetJml(vPrimary).Value > 0 Then
                    MsgBox("There sales order refering to this data, you cannot delete this data", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'tidak bisa dihapus, jika sudah digunakan pada invoice
            If DTA_Inv.ScalarQuery_GetJml(vPrimary).HasValue = True Then
                If DTA_Inv.ScalarQuery_GetJml(vPrimary).Value > 0 Then
                    MsgBox("There sales invoice refering to this data, you cannot delete this data", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'tidak bisa dihapus, jika sudah digunakan pada invoice obselete
            If DTA_Inv_Obs.ScalarQuery_GetJml(vPrimary).HasValue = True Then
                If DTA_Inv_Obs.ScalarQuery_GetJml(vPrimary).Value > 0 Then
                    MsgBox("There obselete invoice refering to this data, you cannot delete this data", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'tidak bisa dihapuskan, jika sudah ada sales budget yang merujuk
            If DTABudget.ScalarQuery_GetJml(vPrimary).HasValue = True Then
                If DTABudget.ScalarQuery_GetJml(vPrimary).Value > 0 Then
                    MsgBox("There sales budget refering to this data, you cannot delete this data", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
            DTA_Cust.Dispose()
            DTA_SO.Dispose()
            DTA_Inv.Dispose()
            DTA_Inv_Obs.Dispose()
            DTABudget.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterSalesPerson.tblmaster_salesperson1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salesperson1TableAdapter
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
            ElseIf SearchCriteria = EnumSearchCriteria.ByGroupCode Then
                QueryResult = DTA.FillBy_GroupCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByGroupDesc Then
                QueryResult = DTA.FillBy_GroupDescLIKE(TBL, Criteria + "%")
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
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + _
                    " miliseconds"
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
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter
        Dim TBL As New DseMasterSalesPerson.tblmaster_salespersonDataTable
        Dim row As DseMasterSalesPerson.tblmaster_salespersonRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(Primary))

            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseMasterSalesPerson.tblmaster_salespersonRow)
                With row
                    Me.mPrimary = row.Primary
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mDescription = .Description
                    Me.mModifiedBy = .Modifed_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mCode = .Code
                    Me.mTimestamp = .Timestamp
                    Me.mPrimaryGroup = .primary_group
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
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter

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

    Public Function GetName(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter

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

    Public Function GetJml(ByVal vCode As String) As Double
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetJml(vCode)) = False Then
                Return DTA.ScalarQuery_GetJml(vCode)
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
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(CDec(vPrimary)).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryGroup(ByVal vPrimarySalesP As Double) As Double
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryGroup(vPrimarySalesP)
            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salespersonTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsikanChecklistbox(ByVal vTBL As DseMasterSalesPerson.tblmaster_salesperson2DataTable, ByVal vClb As CheckedListBox)
        Dim DTA As New DseMasterSalesPersonTableAdapters.tblmaster_salesperson2TableAdapter

        Try
            vClb.DataSource = Nothing
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(vTBL)

            vClb.DataSource = vTBL.DefaultView
            vClb.DisplayMember = "Description"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class
