Public Class ClassHarddrive
    Private mModel As String = ""
    Private mtype As String = ""
    Private mserialNo As String = ""

    Public Property Model() As String
        Get
            Return mModel
        End Get
        Set(ByVal value As String)
            mModel = value
        End Set
    End Property

    Public Property Type() As String
        Get
            Return mtype
        End Get
        Set(ByVal value As String)
            mtype = value
        End Set
    End Property

    Public Property SerialNo() As String
        Get
            Return mserialNo
        End Get
        Set(ByVal value As String)
            mserialNo = value
        End Set
    End Property

End Class
