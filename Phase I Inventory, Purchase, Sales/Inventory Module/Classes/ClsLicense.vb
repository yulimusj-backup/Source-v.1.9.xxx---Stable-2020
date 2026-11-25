Option Explicit On
Option Strict Off

Public Class ClsLicense
    Public Function GetExpiryDate(ByVal vLicenseKey As String) As Date
        Dim d1, d2, m1, m2, y1, y2, y3, y4, MonthStr As String
        Dim a As Date = Today

        Try
            d1 = vLicenseKey.Substring(24, 1)
            d2 = vLicenseKey.Substring(0, 1)
            m1 = vLicenseKey.Substring(6, 1)
            m2 = vLicenseKey.Substring(12, 1)
            y1 = vLicenseKey.Substring(9, 1)
            y2 = vLicenseKey.Substring(21, 1)
            y3 = vLicenseKey.Substring(15, 1)
            y4 = vLicenseKey.Substring(18, 1)

            If m1 + m2 = "01" Then
                MonthStr = "Jan"
            ElseIf m1 + m2 = "02" Then
                MonthStr = "Feb"
            ElseIf m1 + m2 = "03" Then
                MonthStr = "Mar"
            ElseIf m1 + m2 = "04" Then
                MonthStr = "Apr"
            ElseIf m1 + m2 = "05" Then
                MonthStr = "May"
            ElseIf m1 + m2 = "06" Then
                MonthStr = "Jun"
            ElseIf m1 + m2 = "07" Then
                MonthStr = "Jul"
            ElseIf m1 + m2 = "08" Then
                MonthStr = "Aug"
            ElseIf m1 + m2 = "09" Then
                MonthStr = "Sep"
            ElseIf m1 + m2 = "10" Then
                MonthStr = "Oct"
            ElseIf m1 + m2 = "11" Then
                MonthStr = "Nov"
            Else
                MonthStr = "Dec"
            End If

            Return CDate(d1 + d2 + " " + MonthStr + " " + y1 + y2 + y3 + y4)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
