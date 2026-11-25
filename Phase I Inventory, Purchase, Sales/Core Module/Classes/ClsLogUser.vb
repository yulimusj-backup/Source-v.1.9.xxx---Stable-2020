Option Explicit On
Option Strict Off
Public Class ClsLogUser
    Private mPrimary As Double = -1
    Private mActivityType As Integer = 0
    Private mActivityTime As Date
    Private mUser As String = ""
    Private mRemarks As String = ""

    Public Enum EnumActivityType As Integer
        All_Actifity
        Create
        Modify
        Delete
    End Enum

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property ActivityType() As Integer
        Get
            Return Me.mActivityType
        End Get
        Set(ByVal value As Integer)
            mActivityType = value
        End Set
    End Property


    Public Property ActivityTime() As Date
        Get
            Return Me.mActivityTime
        End Get
        Set(ByVal value As Date)
            mActivityTime = value
        End Set
    End Property

    Public Property User() As String
        Get
            Return Me.mUser
        End Get
        Set(ByVal value As String)
            mUser = value
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

End Class
