Option Explicit On
Option Strict Off

Imports System.Management

Public Class ClsLicense
    Dim searcher As New Management.ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive")
    Dim searcher2 As New Management.ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia")
    Dim hd As New ClassHarddrive
    Dim wmi_HD As Management.ManagementObject
    Dim HdCollection As New Collection

    Dim strMotherBoardID As String = String.Empty
    Dim strProcessorId As String = String.Empty
    Dim MACAddress As String = String.Empty
    Dim VolumeSerial As String = String.Empty

    Dim tmpStr2 As String = ""
    Dim myScop As New Management.ManagementScope("\\" & Environment.MachineName & "\root\cimv2")
    Dim oQuer As New Management.SelectQuery("SELECT * FROM WIN32_DiskDrive")
    Dim oResult As New Management.ManagementObjectSearcher(myScop, oQuer)
    Dim oIte As Management.ManagementObject
    Dim oPropert As Management.PropertyData

    Public Enum EnumTipeLicense As Integer
        Trial
        Normal1
        Normal2
        Normal3
        Normal4
        Normal5
        Normal6
        Normal7
        Normal8
        Normal9
        Normal10
    End Enum

    Private Function GetHardDiskSerialNumber() As String
        Try
            ''Langkah 1
            For Each wmi_HD In searcher.Get
                hd.Model = wmi_HD("Model").ToString
                hd.Type = wmi_HD("InterfaceType").ToString
                HdCollection.Add(hd)
                Exit For
            Next

            ''Langkah 2
            For i As Integer = 1 To searcher2.Get.Count
                ''get the hard drive from collection
                ''using index
                hd = TryCast(HdCollection.Item(i), ClassHarddrive)
                ''get the hardware serial no.
                If String.IsNullOrEmpty(wmi_HD("SerialNumber").ToString) = True Then
                    hd.SerialNo = "None"
                Else
                    hd.SerialNo = wmi_HD("SerialNumber").ToString
                End If
                Exit For
            Next

            ''Langkah 3
            If HdCollection.Count > 0 Then
                hd = TryCast(HdCollection.Item(1), ClassHarddrive)
                Return hd.SerialNo
            Else
                Return "None"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "None"
        End Try
    End Function

    Private Function GetMBSerial() As String
        Dim query As New SelectQuery("Win32_BaseBoard")
        Dim search As New ManagementObjectSearcher(query)
        Dim info As ManagementObject

        Try
            For Each info In search.Get()
                strMotherBoardID = info("SerialNumber").ToString()
            Next
            Return strMotherBoardID
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "None"
        End Try
    End Function

    Private Function GetProcessorSerial() As String
        Dim query As New SelectQuery("Win32_processor")
        Dim search As New ManagementObjectSearcher(query)
        Dim info As ManagementObject
        Try
            For Each info In search.Get()
                strProcessorId = info("processorId").ToString()
            Next
            Return strProcessorId
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "None"
        End Try
    End Function

    Private Function Enkripsi1(ByVal PasswordUntukDiEnkripsi As String, Optional ByVal Salt As String = "") As String
        Try
            Dim SHA1 As System.Security.Cryptography.SHA1CryptoServiceProvider
            Dim bytValue() As Byte
            Dim bytHash() As Byte

            ' Create New Crypto Service Provider Object
            SHA1 = New System.Security.Cryptography.SHA1CryptoServiceProvider

            ' Convert the original string to array of Bytes
            bytValue = _
             System.Text.Encoding.UTF8.GetBytes(PasswordUntukDiEnkripsi + Salt)

            ' Compute the Hash, returns an array of Bytes
            bytHash = SHA1.ComputeHash(bytValue)

            SHA1.Clear()

            ' Return a base 64 encoded string of the Hash value
            Return Convert.ToBase64String(bytHash)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try

    End Function

    Private Function Enkripsi2(ByVal PasswordUntukDiEnkripsi As String, Optional ByVal Salt As String = "") As String
        Try
            Dim SHA1 As System.Security.Cryptography.MD5CryptoServiceProvider
            Dim bytValue() As Byte
            Dim bytHash() As Byte

            ' Create New Crypto Service Provider Object
            SHA1 = New System.Security.Cryptography.MD5CryptoServiceProvider

            ' Convert the original string to array of Bytes
            bytValue = _
             System.Text.Encoding.UTF8.GetBytes(PasswordUntukDiEnkripsi + Salt)

            ' Compute the Hash, returns an array of Bytes
            bytHash = SHA1.ComputeHash(bytValue)

            SHA1.Clear()

            ' Return a base 64 encoded string of the Hash value
            Return Convert.ToBase64String(bytHash)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan license key berdasarkan product key, tipe license, tanggal start license
    ''' terdiri dari 5 set data
    ''' 1. Tipe License
    ''' 2. Tanggal Expired
    ''' 3. Hubungan dengan product key
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLicenseKey(ByVal vTanggalStartLicense As Date, ByVal vProductKey As String, _
    ByVal vTipeLicense As EnumTipeLicense) As String
        Dim EndDate As Date
        Dim DateStr, MonthStr, YearStr, d1, d2, m1, m2, y1, y2, y3, y4, T As String
        Dim Dummy1, Dummy2 As String
        Dim Filler As New List(Of String)

        Try
            'persiapan expiry date
            If vTipeLicense = EnumTipeLicense.Trial Then
                EndDate = vTanggalStartLicense + TimeSpan.FromDays(35)
                T = "0"
            ElseIf vTipeLicense = EnumTipeLicense.Normal1 Then
                EndDate = vTanggalStartLicense + TimeSpan.FromDays(370)
                T = "1"
            ElseIf vTipeLicense = EnumTipeLicense.Normal2 Then
                EndDate = vTanggalStartLicense + TimeSpan.FromDays(370)
                T = "2"
            ElseIf vTipeLicense = EnumTipeLicense.Normal3 Then
                EndDate = vTanggalStartLicense + TimeSpan.FromDays(370)
                T = "3"
            ElseIf vTipeLicense = EnumTipeLicense.Normal4 Then
                EndDate = vTanggalStartLicense + TimeSpan.FromDays(370)
                T = "4"
            ElseIf vTipeLicense = EnumTipeLicense.Normal5 Then
                EndDate = vTanggalStartLicense + TimeSpan.FromDays(370)
                T = "5"
            ElseIf vTipeLicense = EnumTipeLicense.Normal6 Then
                EndDate = vTanggalStartLicense + TimeSpan.FromDays(370)
                T = "6"
            ElseIf vTipeLicense = EnumTipeLicense.Normal7 Then
                EndDate = vTanggalStartLicense + TimeSpan.FromDays(370)
                T = "7"
            ElseIf vTipeLicense = EnumTipeLicense.Normal8 Then
                EndDate = vTanggalStartLicense + TimeSpan.FromDays(370)
                T = "8"
            Else
                EndDate = vTanggalStartLicense + TimeSpan.FromDays(370)
                T = "9"
            End If

            DateStr = EndDate.Day.ToString
            If DateStr.Length < 2 Then
                DateStr = "0" + DateStr
            End If
            MonthStr = EndDate.Month.ToString
            If MonthStr.Length < 2 Then
                MonthStr = "0" + MonthStr
            End If
            YearStr = EndDate.Year.ToString
            d1 = DateStr.Substring(0, 1)
            d2 = DateStr.Substring(1)
            m1 = MonthStr.Substring(0, 1)
            m2 = MonthStr.Substring(1)
            y1 = YearStr.Substring(0, 1)
            y2 = YearStr.Substring(1, 1)
            y3 = YearStr.Substring(2, 1)
            y4 = YearStr.Substring(3, 1)

            'pecah product key menjadi 2 group
            vProductKey.Replace(" ", "")
            'isikan fillter
            If vTipeLicense = EnumTipeLicense.Trial Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)))
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)))
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal1 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "1")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "1")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal2 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "2")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "2")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal3 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "3")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "3")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal4 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "4")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "4")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal5 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "5")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "5")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal6 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "6")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "6")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal7 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "7")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "7")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal8 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "8")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "8")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal9 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "9")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "9")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            Else : vTipeLicense = EnumTipeLicense.Normal10
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "10")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "10")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            End If

            Return d2 + Filler(0) + T + Filler(1) + m1 + Filler(2) + y1 + Filler(3) + m2 + Filler(4) + _
                y3 + Filler(5) + y4 + Filler(6) + y2 + Filler(7) + d1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Filler = Nothing
        End Try
    End Function

    'untuk license key yang periode licensenya tidak standar
    Public Function GetLicenseKey(ByVal vTanggalStartLicense As Date, ByVal vProductKey As String, _
    ByVal vTipeLicense As EnumTipeLicense, ByVal EndDate As Date) As String
        Dim DateStr, MonthStr, YearStr, d1, d2, m1, m2, y1, y2, y3, y4, T As String
        Dim Dummy1, Dummy2 As String
        Dim Filler As New List(Of String)

        Try
            'persiapan expiry date
            If vTipeLicense = EnumTipeLicense.Trial Then
                T = "0"
            ElseIf vTipeLicense = EnumTipeLicense.Normal1 Then
                T = "1"
            ElseIf vTipeLicense = EnumTipeLicense.Normal2 Then
                T = "2"
            ElseIf vTipeLicense = EnumTipeLicense.Normal3 Then
                T = "3"
            ElseIf vTipeLicense = EnumTipeLicense.Normal4 Then
                T = "4"
            ElseIf vTipeLicense = EnumTipeLicense.Normal5 Then
                T = "5"
            ElseIf vTipeLicense = EnumTipeLicense.Normal6 Then
                T = "6"
            ElseIf vTipeLicense = EnumTipeLicense.Normal7 Then
                T = "7"
            ElseIf vTipeLicense = EnumTipeLicense.Normal8 Then
                T = "8"
            Else
                T = "9"
            End If

            DateStr = EndDate.Day.ToString
            If DateStr.Length < 2 Then
                DateStr = "0" + DateStr
            End If
            MonthStr = EndDate.Month.ToString
            If MonthStr.Length < 2 Then
                MonthStr = "0" + MonthStr
            End If
            YearStr = EndDate.Year.ToString
            d1 = DateStr.Substring(0, 1)
            d2 = DateStr.Substring(1)
            m1 = MonthStr.Substring(0, 1)
            m2 = MonthStr.Substring(1)
            y1 = YearStr.Substring(0, 1)
            y2 = YearStr.Substring(1, 1)
            y3 = YearStr.Substring(2, 1)
            y4 = YearStr.Substring(3, 1)

            'pecah product key menjadi 2 group
            vProductKey.Replace(" ", "")
            'isikan fillter
            If vTipeLicense = EnumTipeLicense.Trial Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)))
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)))
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal1 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "1")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "1")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal2 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "2")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "2")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal3 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "3")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "3")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal4 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "4")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "4")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal5 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "5")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "5")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal6 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "6")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "6")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal7 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "7")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "7")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal8 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "8")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "8")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            ElseIf vTipeLicense = EnumTipeLicense.Normal9 Then
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "9")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "9")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            Else : vTipeLicense = EnumTipeLicense.Normal10
                Dummy1 = Me.Enkripsi1(vProductKey.Substring(0, CInt(vProductKey.Length / 2)) + "10")
                Dummy2 = Me.Enkripsi2(vProductKey.Substring(CInt(vProductKey.Length / 2)) + "10")
                For i As Integer = 0 To 3
                    Filler.Add(Dummy1.Substring(i * 2, 2))
                Next
                For i As Integer = 4 To 7
                    Filler.Add(Dummy2.Substring(i * 2, 2))
                Next
            End If

            Return d2 + Filler(0) + T + Filler(1) + m1 + Filler(2) + y1 + Filler(3) + m2 + Filler(4) + _
                y3 + Filler(5) + y4 + Filler(6) + y2 + Filler(7) + d1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Filler = Nothing
        End Try
    End Function

    Public Function GetTipeLicense(ByVal vLicenseKey As String) As EnumTipeLicense
        Try
            Return CInt(vLicenseKey.Substring(3, 1))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

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

    ''' <summary>
    ''' product key adalah rangkaian 5 himpunan angka yang masing2 terdiri dari 5 angka
    ''' 1. 5 angka pertama --> 5 set hard disk I
    ''' 2. 5 angka kedua --> 5 set hard disk II
    ''' 3. 5 angka ketiga --> set mother board I
    ''' 4. 5 angka ke empat --> set mother board II
    ''' 5. 5 angka ke lima --> set prosesor
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetProductKey() As String
        Dim HardDiskSN, ProcessorSN, MBSN As String
        Dim retval As String = ""
        Dim Salt As String = ""

        Try
            HardDiskSN = GetVolumeSerial()
            ProcessorSN = Me.GetProcessorSerial
            MBSN = Me.GetMBSerial

            If HardDiskSN.Length > 0 Then
                If Salt = "" Then
                    Salt = HardDiskSN.Substring(0, 1)
                Else
                    Salt = Salt + HardDiskSN.Substring(0, 1)
                End If
            End If
            If ProcessorSN.Length > 0 Then
                If Salt = "" Then
                    Salt = ProcessorSN.Substring(0, 1)
                Else
                    Salt = Salt + ProcessorSN.Substring(0, 1)
                End If
            End If
            If MBSN.Length > 0 Then
                If Salt = "" Then
                    Salt = MBSN.Substring(0, 1)
                Else
                    Salt = Salt + MBSN.Substring(0, 1)
                End If
            End If
            Salt = Me.Enkripsi1(Salt)

            If IsNumeric(Salt.Substring(0, 1)) = True Then
                retval = Me.Enkripsi1(Me.GetSet_I_HD(HardDiskSN), Salt.Substring(1, 1)).ToUpper.Substring(0, 5) + " - " + Me.GetSet_II_HD(HardDiskSN) + " - " _
                + Me.Enkripsi1(Me.GetSet_II_MB(MBSN), Salt.Substring(1, 1)).ToUpper.Substring(0, 5) + " - " + Me.GetSet_I_MB(MBSN) + " - " + _
                Me.Enkripsi1(Me.GetSet_Processor(ProcessorSN)).ToUpper.Substring(0, 5)
            Else
                retval = Me.GetSet_I_HD(HardDiskSN) + " - " + Me.Enkripsi1(Me.GetSet_II_HD(HardDiskSN), Salt.Substring(1, 1)).ToUpper.Substring(0, 5) + " - " _
                + Me.GetSet_II_MB(MBSN) + " - " + Me.Enkripsi1(Me.GetSet_I_MB(MBSN), Salt.Substring(1, 1)).ToUpper.Substring(0, 5) + " - " + _
                Me.Enkripsi1(Me.GetSet_Processor(ProcessorSN)).ToUpper.Substring(0, 5)
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetVolumeSerial(Optional ByVal strDriveLetter As String = "C") As String
        Try
            Dim disk As ManagementObject = New ManagementObject(String.Format("win32_logicaldisk.deviceid=""{0}:""", strDriveLetter))
            disk.Get()
            Return disk("VolumeSerialNumber").ToString()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetSet_Processor(ByVal Processor_SN As String) As String
        Dim ij As Char()

        Try
            If Processor_SN.Length >= 6 Then
                Processor_SN = Processor_SN.Substring(0, Processor_SN.Length - 1)

                ij = Processor_SN.Substring(Processor_SN.Length - 5).ToCharArray
                Return CStr(ij(3).ToString + ij(4).ToString + ij(2).ToString + _
                ij(1).ToString + ij(0).ToString)
            Else
                Return "12345"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "12345"
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan set pertama dari serial number hard disk
    ''' </summary>
    ''' <param name="HDSN"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSet_I_HD(ByVal HDSN As String) As String
        Try
            If HDSN.Length >= 5 Then
                Return HDSN.Substring(HDSN.Length - 5)
            Else
                Return "12345"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "12345"
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan set kedua dari serial number hard disk
    ''' </summary>
    ''' <param name="HDSN"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSet_II_HD(ByVal HDSN As String) As String
        Try
            If HDSN.Length >= 5 Then
                Return HDSN.Substring(0, 5)
            Else
                Return "1234"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "12345"
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan set pertama dari serial number mother board
    ''' </summary>
    ''' <param name="MBSN"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSet_I_MB(ByVal MBSN As String) As String
        Dim ij As Char()

        Try
            If MBSN.Length >= 5 Then
                ij = MBSN.Substring(0, 5).ToCharArray
                Return CStr(ij(3).ToString + ij(4).ToString + ij(2).ToString + _
                ij(1).ToString + ij(0).ToString)
            Else
                Return "cdeBa"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "cdeBa"
        End Try
    End Function

    Private Function GetSet_II_MB(ByVal MBSN As String) As String
        Dim ij As Char()

        Try
            If MBSN.Length >= 12 Then
                ij = MBSN.Substring(7, 5).ToCharArray
                Return CStr(ij(3).ToString + ij(4).ToString + ij(2).ToString + _
                ij(1).ToString + ij(0).ToString)
            Else
                Return "efHkM"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "efHkM"
        End Try
    End Function
End Class
