Option Explicit On
Option Strict Off

Public Class ClsMasterTipeOtorisasi
    Private mPrimary As Double = -1
    Private mPrimaryTipeApproval As Double = -1
    Private mPrimaryRequestBy As Double = -1
    Private mPrimaryApprovedBy As Double = -1
    Private mReason As String = ""
    Private mPrimaryDokumenAcuan As Double = -1
    Private mNoDocumentHist As String = ""

    Public Enum EnumTipeOtorisasi As Integer
        NoTipe
        SO_Edit
        SO_Delete
        SO_Exceed_Credit_Limit
        SO_Overdue_Invoice
        DO_New
        DO_Exceed_Credit_Limit
        DO_Overdue_Invoice
        SF_Override_DO
    End Enum

    Public Property NoDocumentHist() As String
        Get
            Return Me.mNoDocumentHist
        End Get
        Set(ByVal value As String)
            mNoDocumentHist = value
        End Set
    End Property

    Public Property PrimaryDokumenAcuan() As Double
        Get
            Return Me.mPrimaryDokumenAcuan
        End Get
        Set(ByVal value As Double)
            mPrimaryDokumenAcuan = value
        End Set
    End Property

    Public Property Reason() As String
        Get
            Return Me.mReason
        End Get
        Set(ByVal value As String)
            mReason = value
        End Set
    End Property

    Public Property PrimaryApprovedBy() As Double
        Get
            Return Me.mPrimaryApprovedBy
        End Get
        Set(ByVal value As Double)
            mPrimaryApprovedBy = value
        End Set
    End Property

    Public Property PrimaryRequestBy() As Double
        Get
            Return Me.mPrimaryRequestBy
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryRequestBy = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property PrimaryTipeApproval() As EnumTipeOtorisasi
        Get
            Return CInt(Me.mPrimaryTipeApproval)
        End Get
        Set(ByVal value As EnumTipeOtorisasi)
            Me.mPrimaryTipeApproval = value
        End Set
    End Property

    Public Sub SimpanData()
        Dim DTA As New DseLaporanOtorisasiTransaksiTableAdapters.tbltrans_otorisasi_transaksiTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Insert(Me.mPrimaryApprovedBy, Me.mPrimaryRequestBy, Now, Me.mReason, Me.mPrimaryTipeApproval, _
            Me.mPrimaryDokumenAcuan, Today, Me.mNoDocumentHist)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class
