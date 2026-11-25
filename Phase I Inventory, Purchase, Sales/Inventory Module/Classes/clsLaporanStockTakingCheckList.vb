Option Explicit On
Option Strict Off

Public Class clsLaporanStockTakingCheckList
    Private mPrimaryProduct As Double
    Private mDescription As String
    Private mUnitCode As String
    Private mQtyOnHand As Double
    Private mQtyTaking As Double
    Private mLotNo As String
    Private mExpiryDate As Date
    Private mGroup As String
    Private mDifference As String

    Public Property Difference() As String
        Get
            Return mDifference
        End Get
        Set(ByVal value As String)
            mDifference = value
        End Set
    End Property
    Public Property Group() As String
        Get
            Return mGroup
        End Get
        Set(ByVal value As String)
            mGroup = value
        End Set
    End Property
    Public Property PrimaryProduct() As Double
        Get
            Return mPrimaryProduct
        End Get
        Set(ByVal value As Double)
            mPrimaryProduct = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
        End Set
    End Property

    Public Property UnitCode() As String
        Get
            Return mUnitCode
        End Get
        Set(ByVal value As String)
            mUnitCode = value
        End Set
    End Property

    Public Property QtyONHand() As Double
        Get
            Return mQtyOnHand
        End Get
        Set(ByVal value As Double)
            mQtyOnHand = value
        End Set
    End Property

    Public Property QtyTaking() As Double
        Get
            Return mQtyTaking
        End Get
        Set(ByVal value As Double)
            mQtyTaking = value
        End Set
    End Property

    Public Property LotNO() As String
        Get
            Return mLotNo
        End Get
        Set(ByVal value As String)
            mLotNo = value
        End Set
    End Property

    Public Property ExpiryDate() As Date
        Get
            Return mExpiryDate
        End Get
        Set(ByVal value As Date)
            mExpiryDate = value
        End Set
    End Property
End Class
