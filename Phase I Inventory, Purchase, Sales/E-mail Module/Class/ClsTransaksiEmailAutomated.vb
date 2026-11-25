Option Explicit On
Option Strict Off

''' <summary>
''' kelas ini merupakan kelas yang digunakan untuk setting automated email oleh user.
''' Tidak semua automated email bisa disetting hanya email - email tertentu saja, contoh yang tidak bisa
''' disetting: email yang dikirimkan komputer setelah ada booking yang digugurkan, komputer secara
''' langsung akan mengirimkan email ini ke user yang membuat booking
''' </summary>
''' <remarks></remarks>
Public Class ClsTransaksiEmailAutomated
    Dim mPrimary As Double = -1
    Dim mSubject As String = ""
    Dim mDescription As String = ""

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            Me.mPrimary = value
        End Set
    End Property

    Public Property SubJect() As String
        Get
            Return Me.mSubject
        End Get
        Set(ByVal value As String)
            mSubject = value
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

    Public Function CariDetail(ByVal vPrimary As Double, ByVal Clb As Windows.Forms.CheckedListBox) As Boolean
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_email_automatedTableAdapter
        Dim DTA1 As New DsTransaksiEmailTableAdapters.tbltrans_email_automated_sendtoTableAdapter
        Dim TBL As New DsTransaksiEmail.tbltrans_email_automated_sendtoDataTable
        Dim row As DsTransaksiEmail.tbltrans_email_automated_sendtoRow
        Dim rowUser As Core_Module.DSUser.tbluser1Row
        Dim Dataview As Data.DataRowView

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            'kosongkan check box
            While Clb.CheckedItems.Count > 0
                Clb.SetItemCheckState(Clb.CheckedIndices(0), Windows.Forms.CheckState.Unchecked)
            End While
            Me.mDescription = DTA.ScalarQuery_GetDescription(CDec(vPrimary))
            'isikan tabel send to
            DTA1.FillBy_PrimaryEmail(TBL, CDec(vPrimary))
            Dim idx As Integer = 0
            For Each row In TBL.Rows
                idx = 0
                For Each Dataview In Clb.Items
                    rowUser = TryCast(Dataview.Row, Core_Module.DSUser.tbluser1Row)
                    If rowUser.Kode_User = row.User_Code Then
                        Clb.SetItemCheckState(idx, Windows.Forms.CheckState.Checked)
                        Exit For
                    End If
                    idx += 1
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function SimpanData(ByVal vPrimary As Double, ByVal clb As Windows.Forms.CheckedListBox) As Boolean
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_email_automated_sendtoTableAdapter
        Dim rowUser As Core_Module.DSUser.tbluser1Row
        Dim Dataview As Data.DataRowView

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan data
            DTA.Delete(CDec(vPrimary))

            'masukkan data baru
            For Each Dataview In clb.CheckedItems
                rowUser = TryCast(Dataview.Row, Core_Module.DSUser.tbluser1Row)
                DTA.Insert(vPrimary, rowUser.Kode_User)
            Next
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiListBox(ByVal lbo As Windows.Forms.ListBox, ByVal TBL _
    As DsTransaksiEmail.tbltrans_email_automatedDataTable) As Boolean
        Dim DTA As New DsTransaksiEmailTableAdapters.tbltrans_email_automatedTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL)
            lbo.DataSource = TBL.DefaultView
            lbo.DisplayMember = "Subject English"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
