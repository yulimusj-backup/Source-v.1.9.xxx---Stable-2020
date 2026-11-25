Option Explicit On
Option Strict Off

Public Class ClsTransaksiReminder

    Private mPrimary As Double = -1
    Private mSubject As String = ""
    Private mMemo As String = ""
    Private mTanggal As Date = Now

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

    Public Property Subject() As String
        Get
            Return mSubject
        End Get
        Set(ByVal value As String)
            mSubject = value
        End Set
    End Property

    Public Property Memo() As String
        Get
            Return mMemo
        End Get
        Set(ByVal value As String)
            mMemo = value
        End Set
    End Property

    Public Property Tanggal() As Date
        Get
            Return mTanggal
        End Get
        Set(ByVal value As Date)
            mTanggal = value
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

    Public Sub CariDetail(ByVal vclb As Windows.Forms.CheckedListBox, ByVal vTBL As Core_Module.DSUser.tbluser1DataTable, _
    ByVal VPrimaryReminder As Double)
        Dim rowView As DataRowView
        Dim row As Core_Module.DSUser.tbluser1Row
        Dim DTAItem As New DseTransaksiReminderTableAdapters.tbltrans_reminder_itemTableAdapter
        Dim TBLItem As New DseTransaksiReminder.tbltrans_reminder_itemDataTable
        Dim rowItem As DseTransaksiReminder.tbltrans_reminder_itemRow
        Dim DTAHeader As New DseTransaksiReminderTableAdapters.tbltrans_reminderTableAdapter
        Dim TBLHeader As New DseTransaksiReminder.tbltrans_reminderDataTable
        Dim RowHeader As DseTransaksiReminder.tbltrans_reminderRow
        Try
            'cari data header
            DTAHeader.FillBy_Primary(TBLHeader, VPrimaryReminder)
            If TBLHeader.Rows.Count > 0 Then
                RowHeader = TBLHeader.Rows(0)
                With RowHeader
                    Me.mPrimary = .Primary
                    Me.mSubject = .Subject
                    Me.mMemo = .Memo
                    Me.mTanggal = .Tanggal
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
            End If

            'cari data item
            '0. Load Data item
            DTAItem.FillBy_PrimaryReminder(TBLItem, VPrimaryReminder)
            '1. uncheck semua item
            For ij As Integer = 0 To vclb.Items.Count - 1
                vclb.SetItemCheckState(ij, CheckState.Unchecked)
            Next
            '2 check hanya kepada yang ditujukan
            For Each rowItem In TBLItem.Rows
                For ij As Integer = 0 To vclb.Items.Count - 1
                    rowView = vclb.Items(ij)
                    row = rowView.Row
                    If row.Kode_User = rowItem.Kode_User Then
                        vclb.SetItemCheckState(ij, CheckState.Checked)
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            DTAHeader.Dispose()
            TBLHeader.Dispose()
        End Try
    End Sub

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, _
        Optional ByVal usercode As String = "") As Boolean
        Try
            'data tidak boleh edit jika bukan yang buat
            Dim CreatedByLama As String = Me.GetCreatedBy(vPrimary)
            If CreatedByLama <> usercode Then
                MsgBox("Only user with ID:" + CreatedByLama + " can delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetCreatedBy(ByVal vPrimaryRemider As Double) As String
        Dim DTA As New DseTransaksiReminderTableAdapters.tbltrans_reminderTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimaryRemider)) = False Then
                Return DTA.ScalarQuery_GetCreatedBy(vPrimaryRemider)
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
    Optional ByVal usercode As String = "") As Boolean
        Try
            'data tidak boleh edit jika bukan yang buat
            Dim CreatedByLama As String = Me.GetCreatedBy(vPrimaryTransaksi)
            If CreatedByLama <> usercode Then
                MsgBox("Only user with ID:" + CreatedByLama + " can edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vAccountCode As String) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByAccountCode(vAccountCode).HasValue Then
                If DTA.ScalarQuery_GetJmlByAccountCode(vAccountCode).Value > 0 Then
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

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
            'If vBaru = True Then
            'If Me.ApakahSudahAda(Me.mAccountCode) = True Then
            'MsgBox("Sudah ada akun dengan no " + Me.mAccountCode + ", ganti akun code.", MsgBoxStyle.Information, "Information")
            'Return False
            'End If
            'Else
            'no po tidak boleh duplikat
            'Dim NoLama As String = Me.GetAccountCode(Me.mPrimary)
            'If Me.AccountCode <> NoLama Then
            'If Me.ApakahSudahAda(Me.mAccountCode) = True Then
            'MsgBox("Sudah ada akun dengan no " + Me.mAccountCode + ", ganti akun code.", MsgBoxStyle.Information, "Information")
            'Return False
            'End If
            'End If
            'End If

            'return value
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJml(ByVal vTanggalTransaksi As Date, ByVal vUserName As String) As Double
        Dim DTA As New DseTransaksiReminderTableAdapters.tbltrans_reminderTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml(vTanggalTransaksi, vUserName).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vSubject As String, ByVal vCreatedBy As String, ByVal vCreatedDate As Date) As Double
        Dim DTA As New DseTransaksiReminderTableAdapters.tbltrans_reminderTableAdapter
        Try
            If DTA.ScalarQuery_GetPrimaryBySubject(vSubject, vCreatedBy, vCreatedDate).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryBySubject(vSubject, vCreatedBy, vCreatedDate).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiReminderTableAdapters.tbltrans_reminderTableAdapter
        Try
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanBaru(ByVal vclb As Windows.Forms.CheckedListBox, ByVal vTBL As Core_Module.DSUser.tbluser1DataTable, _
    ByVal vUserCode As String) As Boolean
        Dim DTAItem As New DseTransaksiReminderTableAdapters.tbltrans_reminder_itemTableAdapter
        Dim rowView As DataRowView
        Dim row As Core_Module.DSUser.tbluser1Row
        Dim DTA As New DseTransaksiReminderTableAdapters.tbltrans_reminderTableAdapter
        Dim CreatedDate_Dummy As Date = Now

        Try
            If DTA.Insert(Me.Subject, Me.Memo, _
                        vUserCode, CreatedDate_Dummy, vUserCode, CreatedDate_Dummy, CreatedDate_Dummy, Me.Tanggal) > 0 Then
                Me.mPrimary = Me.GetPrimary(Me.mSubject, vUserCode, CreatedDate_Dummy)
                Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan item
                For Each rowView In vclb.CheckedItems
                    row = rowView.Row
                    DTAItem.Insert(Me.mPrimary, row.Kode_User)
                Next

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vclb As Windows.Forms.CheckedListBox, ByVal vTBL As Core_Module.DSUser.tbluser1DataTable, _
    ByVal vUserCode As String) As Boolean
        Dim DTAItem As New DseTransaksiReminderTableAdapters.tbltrans_reminder_itemTableAdapter
        Dim rowView As DataRowView
        Dim row As Core_Module.DSUser.tbluser1Row
        Dim DTA As New DseTransaksiReminderTableAdapters.tbltrans_reminderTableAdapter
        Try
            If DTA.Update(Me.Subject, Me.Memo, _
               vUserCode, Now, Now, Me.mTanggal, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'hapus item lama
                For Each rowView In vclb.CheckedItems
                    row = rowView.Row
                    DTAItem.Delete(Me.mPrimary)
                Next

                'simpan item baru 
                For Each rowView In vclb.CheckedItems
                    row = rowView.Row
                    DTAItem.Insert(Me.mPrimary, row.Kode_User)
                Next
                ''For ij As Integer = 0 To vclb.Items.Count - 1
                ''rowView = vclb.Items(ij)
                ''row = rowView.Row
                ''DTAItem.Insert(Me.mPrimary, row.Kode_User)
                ''Next
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double, ByVal userCode As String) As Boolean
        Dim DTA As New DseTransaksiReminderTableAdapters.tbltrans_reminderTableAdapter
        Try
            If Me.ApakahBolehHapus(vPrimary, userCode) = True Then
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

    Public Function IsiDGV(ByVal DGV As DataGridView, ByVal TBL As DseTransaksiReminder.tbltrans_reminder1DataTable, _
    ByVal vKodeUser As String, ByVal vTanggal As Date) As Boolean
        Dim DTA As New DseTransaksiReminderTableAdapters.tbltrans_reminder1TableAdapter
        ''Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DTA.Fill(TBL, vKodeUser, vTanggal)

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(3).Visible = False
                DGV.Columns(1).HeaderText = "Subject"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Created By"
                DGV.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
                DGV.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
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
End Class
