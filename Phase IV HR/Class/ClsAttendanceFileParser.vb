Public Class ClsAttendanceFileParser

    Private mLogTime As Date
    Public Property LogTime() As Date
        Get
            Return mLogTime
        End Get
        Set(ByVal value As Date)
            mLogTime = value
        End Set
    End Property

    Private mUserID As String
    Public Property UserID() As String
        Get
            Return mUserID
        End Get
        Set(ByVal value As String)
            mUserID = value
        End Set
    End Property

    Private mYearIndex As Integer = 0
    Public Property YearIndex() As Integer
        Get
            Return mYearIndex
        End Get
        Set(ByVal value As Integer)
            mYearIndex = value
        End Set
    End Property

    Private mYearLength As Integer = 4
    Public Property YearLength() As Integer
        Get
            Return mYearLength
        End Get
        Set(ByVal value As Integer)
            mYearLength = value
        End Set
    End Property

    Private mMonthIndex As Integer = 4
    Public Property MonthIndex() As Integer
        Get
            Return mMonthIndex
        End Get
        Set(ByVal value As Integer)
            mMonthIndex = value
        End Set
    End Property

    Private mMonthLength As Integer = 2
    Public Property MonthLength() As Integer
        Get
            Return mMonthLength
        End Get
        Set(ByVal value As Integer)
            mMonthLength = value
        End Set
    End Property

    Private mDayIndex As Integer = 6
    Public Property DayIndex() As Integer
        Get
            Return mDayIndex
        End Get
        Set(ByVal value As Integer)
            mDayIndex = value
        End Set
    End Property

    Private mDayLength As Integer = 2
    Public Property DayLength() As Integer
        Get
            Return mDayLength
        End Get
        Set(ByVal value As Integer)
            mDayLength = value
        End Set
    End Property

    Private mHourIndex As Integer = 8
    Public Property HourIndex() As Integer
        Get
            Return mHourIndex
        End Get
        Set(ByVal value As Integer)
            mHourIndex = value
        End Set
    End Property

    Private mHourLength As Integer = 2
    Public Property HourLength() As Integer
        Get
            Return mHourLength
        End Get
        Set(ByVal value As Integer)
            mHourLength = value
        End Set
    End Property

    Private mMinuteIndex As Integer = 10
    Public Property MinuteIndex() As Integer
        Get
            Return mMinuteIndex
        End Get
        Set(ByVal value As Integer)
            mMinuteIndex = value
        End Set
    End Property

    Private mMinuteLength As Integer = 2
    Public Property MinuteLength() As Integer
        Get
            Return mMinuteLength
        End Get
        Set(ByVal value As Integer)
            mMinuteLength = value
        End Set
    End Property

    Private mUserIdIndex As Integer = 12
    Public Property UserIdIndex() As Integer
        Get
            Return mUserIdIndex
        End Get
        Set(ByVal value As Integer)
            mUserIdIndex = value
        End Set
    End Property

    Private mUserIdLength As Integer = 6
    Public Property UserIdLength() As Integer
        Get
            Return mUserIdLength
        End Get
        Set(ByVal value As Integer)
            mUserIdLength = value
        End Set
    End Property

    Private mRemoveZero As Boolean = False
    Public Property RemoveZero() As Boolean
        Get
            Return mRemoveZero
        End Get
        Set(ByVal value As Boolean)
            mRemoveZero = value
        End Set
    End Property

    Public Sub SetIndexAndLength(ByVal vYearIndex As Integer, ByVal vYearLength As Integer, ByVal vMonthIndex As Integer, _
            ByVal vMonthLength As Integer, ByVal vDayIndex As Integer, ByVal vDayLength As Integer, _
            ByVal vHourIndex As Integer, ByVal vHourLength As Integer, ByVal vMinuteIndex As Integer, _
            ByVal vMinuteLength As Integer, ByVal vUserIdIndex As Integer, ByVal vUserIdLength As Integer)
        Try
            Me.YearIndex = vYearIndex
            Me.YearLength = vYearLength
            Me.MonthIndex = vMonthIndex
            Me.MonthLength = vMonthLength
            Me.DayIndex = vDayIndex
            Me.DayLength = vDayLength
            Me.HourIndex = vHourIndex
            Me.HourLength = vHourLength
            Me.MinuteIndex = vMinuteIndex
            Me.MinuteLength = vMinuteLength
            Me.UserIdIndex = vUserIdIndex
            Me.UserIdLength = vUserIdLength
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetLog(ByVal FileName As String, Optional ByVal CutZeroBeforeUserID As Boolean = True) As List(Of ClsAttendanceFileParser)
        Dim Content As String
        Dim Lines() As String
        Dim DateTime As String
        Dim Att As New List(Of ClsAttendanceFileParser)
        Try
            Using SR As New System.IO.StreamReader(FileName)
                Content = SR.ReadToEnd
                Lines = Content.Split(vbNewLine)
                For Each Line As String In Lines
                    Line = Line.Trim
                    Try
                        DateTime = CDate(Line.Substring(Me.mYearIndex, Me.mYearLength) & "/" & Line.Substring(Me.mMonthIndex, Me.mMonthLength) & "/" & Line.Substring(Me.mDayIndex, Me.mDayLength) & " " & Line.Substring(Me.mHourIndex, Me.mHourLength) & ":" & Line.Substring(Me.mMinuteIndex, Me.mMinuteLength) & ":00")
                        Dim AttItem As New ClsAttendanceFileParser
                        AttItem.LogTime = DateTime
                        If CutZeroBeforeUserID Then
                            AttItem.UserID = CInt(Line.Substring(Me.mUserIdIndex, Me.mUserIdLength)).ToString
                        Else
                            AttItem.UserID = Line.Substring(Me.mUserIdIndex, Me.mUserIdLength)
                        End If
                        Att.Add(AttItem)
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try
                Next
            End Using
            Return Att
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Att
        End Try
    End Function
End Class
