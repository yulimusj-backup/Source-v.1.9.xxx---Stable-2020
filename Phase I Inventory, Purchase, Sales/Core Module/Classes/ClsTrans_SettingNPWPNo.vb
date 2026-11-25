Option Explicit On
Option Strict Off

Public Class ClsTrans_SettingNPWPNo

    'Public TglBerlaku As Date = CDate("2023-11-1")

    Public Enum TipeNPWP
        Company
        Individual
    End Enum

    Public Enum EnumFromMaster
        Customer
        Supplier
        Karyawan
    End Enum

    Public Enum EnumNPWPDigit
        _15Digit
        _16Digit
    End Enum

    '2025 Core Tax Tahap 2.1
    '//TR : Tambahan baru Nik/Pas
    Public Enum EnumOptionNikPass As Integer
        ByNIK
        ByPASSPORT
        ByOtherID
    End Enum

    Public Function Tglberlaku() As Date
        Dim objSetting As New ClsTransaksiSetting
        Dim dummy As String = ""
        Try
            dummy = objSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.NPWP, ClsTransaksiSetting.EnumNamaParameter.NPWP_StartPeriode)
            If dummy = "" Then
                Return CDate("7/1/2024")
            Else
                Return CDate(dummy)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'tambah tanggal strict
    Public Function TglberlakuStrict() As Date
        Dim objSetting As New ClsTransaksiSetting
        Dim dummy As String = ""
        Try
            dummy = objSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.NPWP, ClsTransaksiSetting.EnumNamaParameter.NPWP_StartPeriode)
            If dummy = "" Then
                Return Tglberlaku()
            Else
                Return CDate(dummy)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Function ApakahDataSudahBenar_MasterCustomer(ByVal vNPWPNo As String, ByVal vNIKOrPassword As String, ByVal vNPWPLama As String, _
    ByVal vCboNPWPType As ComboBox, ByVal vTampilkanNotif As Boolean) As Boolean

        Try

            If Me.ApakahSudah16Digit(vNPWPNo, True) = False Then
                Return False
            End If

            If vCboNPWPType.SelectedIndex = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                If Me.ApakahNPWPNoDanNIKSama(vNPWPNo, vNIKOrPassword, vTampilkanNotif) = False Then
                    Return False
                End If
            End If

            If String.IsNullOrEmpty(vNPWPLama) = True Then
                MsgBox("Old NPWP No cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If vCboNPWPType.SelectedIndex = -1 Then
                MsgBox("NPWP Type cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    '2025 Core Tax Tahap 2.1
    Public Function ApakahDataSudahBenar_MasterCustomer_New(ByVal vNPWPNo As String, ByVal vNIKOrPassword As String, ByVal vNPWPLama As String, _
    ByVal vCboNPWPType As ComboBox, ByVal vCboNIKType As ComboBox, ByVal vTampilkanNotif As Boolean) As Boolean

        Try

            If vCboNPWPType.SelectedIndex = ClsTrans_SettingNPWPNo.TipeNPWP.Company Then
                If Me.ApakahSudah16Digit_New(vNPWPNo, True) = False Then 'cek format spasi
                    Return False
                End If
            Else
                If Me.ApakahSudah16Digit(vNPWPNo, True) = False Then 'cek length 16 digit
                    Return False
                End If
            End If

            If vCboNPWPType.SelectedIndex = ClsTrans_SettingNPWPNo.TipeNPWP.Individual And vCboNIKType.SelectedIndex = EnumOptionNikPass.ByNIK Then
                If Me.ApakahNPWPNoDanNIKSama(vNPWPNo, vNIKOrPassword, vTampilkanNotif) = False Then
                    Return False
                End If
            ElseIf vCboNPWPType.SelectedIndex = ClsTrans_SettingNPWPNo.TipeNPWP.Individual And vCboNIKType.SelectedIndex <> EnumOptionNikPass.ByNIK Then
                'cek kosong nikpass
                If String.IsNullOrEmpty(vNIKOrPassword.Trim) Then
                    MsgBox("Passport or Other ID cannot empty string.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If String.IsNullOrEmpty(vNPWPLama) = True Then
                MsgBox("Old NPWP No cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If vCboNPWPType.SelectedIndex = -1 Then
                MsgBox("NPWP Type cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahDataSudahBenar_MasterSupplier(ByVal vNPWPNo As String, ByVal vNPWPLama As String, _
  ByVal vCboNPWPType As ComboBox, ByVal vTampilkanNotif As Boolean) As Boolean

        Try

            If Me.ApakahSudah16Digit(vNPWPNo, vTampilkanNotif) = False Then
                Return False
            End If

            If String.IsNullOrEmpty(vNPWPLama) = True Then
                MsgBox("Old NPWP No cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If vCboNPWPType.SelectedIndex = -1 Then
                MsgBox("NPWP Type cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    '2025 Core Tax Tahap 2.1
    Public Function ApakahDataSudahBenar_MasterSupplier_new(ByVal vNPWPNo As String, ByVal vNPWPLama As String, _
ByVal vCboNPWPType As ComboBox, ByVal vTampilkanNotif As Boolean) As Boolean

        Try

            If Me.ApakahSudah16Digit_New(vNPWPNo, vTampilkanNotif) = False Then
                Return False
            End If

            If String.IsNullOrEmpty(vNPWPLama) = True Then
                MsgBox("Old NPWP No cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If vCboNPWPType.SelectedIndex = -1 Then
                MsgBox("NPWP Type cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Function ApakahDataSudahBenar_MasterCompany(ByVal vNPWPNo As String, ByVal vNPWPLama As String, _
 ByVal vTampilkanNotif As Boolean) As Boolean

        Try

            If Me.ApakahSudah16Digit(vNPWPNo, vTampilkanNotif) = False Then
                Return False
            End If

            If String.IsNullOrEmpty(vNPWPLama) = True Then
                MsgBox("Old NPWP No cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    '2025 Core Tax Tahap 2.1
    Public Function ApakahDataSudahBenar_MasterCompany_new(ByVal vNPWPNo As String, ByVal vNPWPLama As String, _
ByVal vTampilkanNotif As Boolean) As Boolean

        Try

            If Me.ApakahSudah16Digit_New(vNPWPNo, vTampilkanNotif) = False Then
                Return False
            End If

            If String.IsNullOrEmpty(vNPWPLama) = True Then
                MsgBox("Old NPWP No cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Function ApakahDataSudahBenar_MasterKaryawan(ByVal vNPWPNo As String, ByVal vNIKOrPassword As String, ByVal vNPWPLama As String, _
    ByVal vTampilkanNotif As Boolean) As Boolean

        Try
            If Me.ApakahSudah16Digit(vNPWPNo, vTampilkanNotif) = False Then
                Return False
            End If

            If Me.ApakahNPWPNoDanNIKSama(vNPWPNo, vNIKOrPassword, False) = False Then
                MsgBox("NPWP No and ID Card must be identical", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If String.IsNullOrEmpty(vNPWPLama) = True Then
                MsgBox("Old NPWP No cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    '2025 Core Tax Tahap 2.1
    Public Function ApakahDataSudahBenar_MasterKaryawan_New(ByVal vNPWPNo As String, ByVal vNIKOrPassword As String, ByVal vNPWPLama As String, _
    ByVal vTampilkanNotif As Boolean) As Boolean

        Try
            If Me.ApakahSudah16Digit_New(vNPWPNo, vTampilkanNotif) = False Then
                Return False
            End If

            If Me.ApakahNPWPNoDanNIKSama(vNPWPNo, vNIKOrPassword, False) = False Then
                MsgBox("NPWP No and ID Card must be identical", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If String.IsNullOrEmpty(vNPWPLama) = True Then
                MsgBox("Old NPWP No cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSudah16Digit(ByVal vNPWPNo As String, ByVal vTampilkanNotif As Boolean) As Boolean
        Dim Retval As String = ""
        Try
            If String.IsNullOrEmpty(vNPWPNo) = False Then
                Retval = vNPWPNo.ToString.Replace(".", "").Replace("-", "").Replace(" ", "")
                If Retval.Length = 16 Then
                    Return True
                End If
            End If

            If vTampilkanNotif = True Then
                MsgBox("Used 16 digit for NPWP No.", MsgBoxStyle.Information, "Information")
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    '2025 Core Tax Tahap 2.1
    Public Function ApakahSudah16Digit_New(ByVal vNPWPNo As String, ByVal vTampilkanNotif As Boolean) As Boolean
        Dim Reference As String = "0000 0000 0000 0000"
        Dim Retval As String = ""
        Dim isFormat As Boolean = True
        Try
            If String.IsNullOrEmpty(vNPWPNo) = False And vNPWPNo.Length = Reference.Length Then
                If vNPWPNo.Contains("-") = True Or vNPWPNo.Contains(".") = True Then
                    isFormat = False
                End If

                'cek 4 digit pertama apakah angka
                Retval = vNPWPNo.Substring(0, 4)
                If IsNumeric(Retval) = False Then
                    isFormat = False
                End If

                'cek spasi pertama
                Retval = vNPWPNo.Substring(4, 1)
                If Retval.Contains(" ") = False Then
                    isFormat = False
                End If

                'cek 4 digit kedua apakah angka
                Retval = vNPWPNo.Substring(5, 4)
                If IsNumeric(Retval) = False Then
                    isFormat = False
                End If

                'cek spasi kedua
                Retval = vNPWPNo.Substring(9, 1)
                If Retval.Contains(" ") = False Then
                    isFormat = False
                End If

                'cek 4 digit ketiga apakah angka
                Retval = vNPWPNo.Substring(10, 4)
                If IsNumeric(Retval) = False Then
                    isFormat = False
                End If

                'cek spasi ketiga
                Retval = vNPWPNo.Substring(14, 1)
                If Retval.Contains(" ") = False Then
                    isFormat = False
                End If

                'cek 4 digit keempat apakah angka
                Retval = vNPWPNo.Substring(15, 4)
                If IsNumeric(Retval) = False Then
                    isFormat = False
                End If
            Else
                isFormat = False
            End If

            If vTampilkanNotif = True And isFormat = False Then
                MsgBox("Used 16 digit format for NPWP No. (ex: 0000 0000 0000 0000)", MsgBoxStyle.Information, "Information")
                Return False
            ElseIf isFormat = False Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSudah15Digit(ByVal vNPWPNo As String, ByVal vTampilkanNotif As Boolean) As Boolean
        Dim Retval As String = ""
        Try
            If String.IsNullOrEmpty(vNPWPNo) = False Then
                Retval = vNPWPNo.ToString.Replace(".", "").Replace("-", "")
                If Retval.Length = 15 Then
                    Return True
                End If
            End If
            If vTampilkanNotif = True Then
                MsgBox("Used 15 digit for NPWP No.", MsgBoxStyle.Information, "Information")
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahNPWPNoDanNIKSama(ByVal vNPWPNo As String, ByVal vNIKOrPassword As String, ByVal vTampilkanNotif As Boolean) As Boolean
        Try
            If String.IsNullOrEmpty(vNPWPNo.Trim) = False And String.IsNullOrEmpty(vNIKOrPassword.Trim) = False Then
                If vNPWPNo = vNIKOrPassword Then
                    Return True
                End If
            End If
            If vTampilkanNotif = True Then
                MsgBox("NPWP No and NIK must be identical", MsgBoxStyle.Information, "Information")
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function UpdateNPWP_MasterCustomer(ByVal vNPWPNo As String, ByVal vNPWPNoOld As String, _
    ByVal vNPWPType As ClsTrans_SettingNPWPNo.TipeNPWP, ByVal vPrimaryCustomer As Double) As Boolean
        'fungis ini mengupdate NPWP, NPWP Lama dan NPWP Type pd Master Customer.
        Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_customerTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.UpdateQuery_CustomerData(vNPWPNo, vNPWPNoOld, vNPWPType, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function UpdateNPWP_MasterCustomer(ByVal vNPWPNo As String, ByVal vNPWPNoOld As String, _
    ByVal vNPWPType As ClsTrans_SettingNPWPNo.TipeNPWP, ByVal vNIK As String, ByVal vPrimaryCustomer As Double) As Boolean
        'fungis ini mengupdate NPWP, NPWP Lama dan NPWP Type pd Master Customer.
        Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_customerTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.UpdateQuery_CustomerData2(vNPWPNo, vNPWPNoOld, vNPWPType, vNIK, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function UpdateNPWP_MasterCustomer(ByVal vNPWPNoOld As String, _
ByVal vNPWPType As ClsTrans_SettingNPWPNo.TipeNPWP, ByVal vPrimaryCustomer As Double) As Boolean
        'fungis ini mengupdate NPWP, NPWP Lama dan NPWP Type pd Master Customer.
        Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_customerTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.UpdateQuery_CustomerData3(vNPWPNoOld, vNPWPType, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function UpdateNPWP_MasterKaryawan(ByVal vNPWPNo As String, ByVal vNPWPNoOld As String, _
    ByVal vNPWPType As ClsTrans_SettingNPWPNo.TipeNPWP, ByVal vNIK As String, ByVal vPrimaryKaryawan As Double) As Boolean
        'fungis ini mengupdate NPWP, NPWP Lama dan NPWP Type pd Master Customer.
        Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_karyawanTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.UpdateQuery_Karyawan2(vNPWPNo, vNPWPNoOld, vNPWPType, vNIK, vPrimaryKaryawan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function UpdateNPWP_MasterKaryawan(ByVal vNPWPNoOld As String, _
   ByVal vNPWPType As ClsTrans_SettingNPWPNo.TipeNPWP, ByVal vPrimaryKaryawan As Double) As Boolean
        'fungis ini mengupdate NPWP, NPWP Lama dan NPWP Type pd Master Customer.
        Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_karyawanTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.UpdateQuery_Karyawan3(vNPWPNoOld, vNPWPType, vPrimaryKaryawan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function UpdateNPWP_MasterSupplier(ByVal vNPWPNo As String, ByVal vNPWPNoOld As String, _
    ByVal vNPWPType As ClsTrans_SettingNPWPNo.TipeNPWP, ByVal vPrimarySupplier As Double) As Boolean
        'fungis ini mengupdate NPWP, NPWP Lama dan NPWP Type pd Master Supplier.
        Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_supplierTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.UpdateQuery_SupplierData(vNPWPNo, vNPWPNoOld, vNPWPType, vPrimarySupplier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function UpdateNPWP_MasterSupplier(ByVal vNPWPNoOld As String, _
   ByVal vNPWPType As ClsTrans_SettingNPWPNo.TipeNPWP, ByVal vPrimarySupplier As Double) As Boolean
        'fungis ini mengupdate NPWP, NPWP Lama dan NPWP Type pd Master Supplier.
        Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_supplierTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.UpdateQuery_SupplierData2(vNPWPNoOld, vNPWPType, vPrimarySupplier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNPWPType_MasterCustomer(ByVal vPrimaryCustomer As Double) As ClsTrans_SettingNPWPNo.TipeNPWP
        Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_customerTableAdapter
        Dim Type As New Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Type = DTA.ScalarQuery_GetNPWPType(vPrimaryCustomer)

            If Type.HasValue = True Then
                Return Type.Value
            End If
            Return TipeNPWP.Company
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNPWPType_MasterSupplier(ByVal vPrimarySupplier As Double) As ClsTrans_SettingNPWPNo.TipeNPWP
        Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_supplierTableAdapter
        Dim Type As New Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Type = DTA.ScalarQuery_GetNPWPType(vPrimarySupplier)

            If Type.HasValue = True Then
                Return Type.Value
            End If
            Return TipeNPWP.Company
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SetNPWP(ByVal vTglTransaksi As Date, ByVal vPrimaryAcuan As Double, ByRef vNPWPNo As String, _
    ByVal vTipeFormMaster As ClsTrans_SettingNPWPNo.EnumFromMaster, Optional ByVal NIK As String = "") As Boolean
        Dim frmNPWP As New Core_Module.FrmTrans_SettingNPWP
        'tambahan setting
        Dim dummy As String = ""
        Dim objSetting As New ClsTransaksiSetting
        Try
            If vTipeFormMaster = EnumFromMaster.Customer Then
                If Me.GetNPWPType_MasterCustomer(vPrimaryAcuan) = ClsTrans_SettingNPWPNo.TipeNPWP.Company Then
                    frmNPWP.typeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Company
                Else
                    frmNPWP.typeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Individual
                End If
            ElseIf vTipeFormMaster = EnumFromMaster.Supplier Then
                If Me.GetNPWPType_MasterSupplier(vPrimaryAcuan) = ClsTrans_SettingNPWPNo.TipeNPWP.Company Then
                    frmNPWP.typeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Company
                Else
                    frmNPWP.typeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Individual
                End If
            ElseIf vTipeFormMaster = EnumFromMaster.Karyawan Then
                frmNPWP.typeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Individual
                frmNPWP.TxtNIK.Text = NIK
            End If

            If vTglTransaksi >= Me.Tglberlaku Then 'Berlaku NPWP 16 digit jika di atas /= 2024-1-1
                frmNPWP.IsNPWPDigit = EnumNPWPDigit._16Digit
                If Me.ApakahSudah16Digit(vNPWPNo, False) = False Then
                    frmNPWP.vTipeFormMaster = vTipeFormMaster
                    frmNPWP.TxtNPWPLama.Text = vNPWPNo
                    frmNPWP.vPrimaryAcuan = vPrimaryAcuan
                    If frmNPWP.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                        Return False
                    End If
                    vNPWPNo = frmNPWP.TxtNPWPNo.Text
                End If
            Else
                frmNPWP.IsNPWPDigit = EnumNPWPDigit._15Digit

                If Me.ApakahSudah15Digit(vNPWPNo, False) = False Then
                    frmNPWP.vTipeFormMaster = vTipeFormMaster
                    frmNPWP.TxtNPWPLama.Text = vNPWPNo
                    frmNPWP.vPrimaryAcuan = vPrimaryAcuan
                    If frmNPWP.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                        Return False
                    End If
                    vNPWPNo = frmNPWP.TxtNPWPNo.Text
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SetNPWP_New(ByVal vTglTransaksi As Date, ByVal vPrimaryAcuan As Double, ByRef vNPWPNo As String, _
        ByVal vTipeFormMaster As ClsTrans_SettingNPWPNo.EnumFromMaster, ByVal vTypeTransaction As frmLaporanVAT2025_EditCustomer.EnumTransactionType, _
        ByVal vUserCode As String, Optional ByVal vPrimaryTransaksi As Double = -1) As Boolean
        Dim frmNPWP_Supl As New Core_Module.frmLaporanVAT2025_EditSupplier
        'tambahan setting
        Dim dummy As String = ""
        Dim objSetting As New ClsTransaksiSetting
        Try
            If vTipeFormMaster = EnumFromMaster.Supplier Then
                If Me.GetNPWPType_MasterSupplier(vPrimaryAcuan) = ClsTrans_SettingNPWPNo.TipeNPWP.Company Then
                    frmNPWP_Supl.cboTypeNPWP.SelectedIndex = ClsTrans_SettingNPWPNo.TipeNPWP.Company
                Else
                    frmNPWP_Supl.cboTypeNPWP.SelectedIndex = ClsTrans_SettingNPWPNo.TipeNPWP.Individual
                End If
            ElseIf vTipeFormMaster = EnumFromMaster.Karyawan Then
                frmNPWP_Supl.cboTypeNPWP.SelectedIndex = ClsTrans_SettingNPWPNo.TipeNPWP.Individual
            End If

            'If vTipeFormMaster <> EnumFromMaster.Supplier Then
            If vTglTransaksi >= Me.Tglberlaku And vTglTransaksi >= Me.TglberlakuStrict Then 'Berlaku NPWP 16 digit jika di atas /= 2024-1-1
                If Me.ApakahSudah16Digit_New(vNPWPNo, False) = False Then
                    frmNPWP_Supl.UserCode = vUserCode
                    frmNPWP_Supl.BackColorAll = Core_Module.GetSMESSBackColorAll
                    frmNPWP_Supl.PrimaryAcuan = vPrimaryAcuan
                    frmNPWP_Supl.PrimaryTransaksi = vPrimaryTransaksi
                    frmNPWP_Supl.TypeTransaction = vTypeTransaction
                    If frmNPWP_Supl.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                        Return False
                    End If
                    vNPWPNo = frmNPWP_Supl.MTxtNPWP.Text
                End If
            ElseIf vTglTransaksi < Me.Tglberlaku And vTglTransaksi < Me.TglberlakuStrict Then
                If Me.ApakahSudah15Digit(vNPWPNo, False) = False Then
                    frmNPWP_Supl.UserCode = vUserCode
                    frmNPWP_Supl.BackColorAll = Core_Module.GetSMESSBackColorAll
                    frmNPWP_Supl.PrimaryAcuan = vPrimaryAcuan
                    frmNPWP_Supl.PrimaryTransaksi = vPrimaryTransaksi
                    frmNPWP_Supl.TypeTransaction = vTypeTransaction
                    If frmNPWP_Supl.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                        Return False
                    End If
                    vNPWPNo = frmNPWP_Supl.MTxtNPWP.Text
                End If
            End If
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SetNPWP_New(ByVal vTglTransaksi As Date, ByVal vPrimaryAcuan As Double, ByRef vNPWPNo As String, _
    ByVal vTipeFormMaster As ClsTrans_SettingNPWPNo.EnumFromMaster, ByRef NIK As String, _
    ByVal vTypeTransaction As frmLaporanVAT2025_EditCustomer.EnumTransactionType, ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String) As Boolean
        Dim frmNPWP_Cust As New Core_Module.frmLaporanVAT2025_EditCustomer
        Dim frmNPWP_Supl As New Core_Module.FrmTrans_SettingNPWP
        'tambahan setting
        Dim dummy As String = ""
        Dim objSetting As New ClsTransaksiSetting
        Try
            If vTipeFormMaster = EnumFromMaster.Customer Then
                If Me.GetNPWPType_MasterCustomer(vPrimaryAcuan) = ClsTrans_SettingNPWPNo.TipeNPWP.Company Then
                    frmNPWP_Cust.cboTypeNPWP.SelectedIndex = ClsTrans_SettingNPWPNo.TipeNPWP.Company
                Else
                    frmNPWP_Cust.cboTypeNPWP.SelectedIndex = ClsTrans_SettingNPWPNo.TipeNPWP.Individual
                End If
            End If

            If vTipeFormMaster = EnumFromMaster.Customer Then
                If vTglTransaksi >= Me.Tglberlaku Then 'Berlaku NPWP 16 digit jika di atas /= 2024-1-1
                    If Me.ApakahSudah16Digit_New(vNPWPNo, False) = False Then
                        frmNPWP_Cust.UserCode = vUserCode
                        frmNPWP_Cust.BackColorAll = Core_Module.GetSMESSBackColorAll
                        frmNPWP_Cust.PrimaryCustomer = vPrimaryAcuan
                        If vTypeTransaction = frmLaporanVAT2025_EditCustomer.EnumTransactionType.TransactionSales Then
                            frmNPWP_Cust.PrimaryInvoice = vPrimaryTransaksi
                        ElseIf vTypeTransaction = frmLaporanVAT2025_EditCustomer.EnumTransactionType.TransactionSalesReturn Then
                            frmNPWP_Cust.PrimaryInvoice = vPrimaryTransaksi
                        End If
                        frmNPWP_Cust.TypeTransaction = vTypeTransaction
                        If frmNPWP_Cust.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                            Return False
                        End If
                        vNPWPNo = frmNPWP_Cust.MTxtNPWP.Text
                        If frmNPWP_Cust.cboTypeNPWP.SelectedIndex = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                            NIK = frmNPWP_Cust.txtNIKPASS.Text
                        End If
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function format16DigitSpasi(ByVal vNPWPNo As String) As String
        Dim Reference As String = "0000 0000 0000 0000"
        Dim dymmyNPWP As String = ""
        Dim Retval As String = ""
        Dim isFormat As Boolean = True
        Try
            If String.IsNullOrEmpty(vNPWPNo) = False And vNPWPNo.Replace(" ", "").Replace("-", "").Replace(".", "").Length = 16 Then
                dymmyNPWP = vNPWPNo.Replace(" ", "").Replace("-", "").Replace(".", "")
                '4 digit pertama
                Retval = dymmyNPWP.Substring(0, 4)
                'tambah spasi pertama
                Retval &= " " & dymmyNPWP.Substring(4, 4)
                'tambah spasi kedua
                Retval &= " " & dymmyNPWP.Substring(8, 4)
                'cek 4 digit ketiga apakah angka
                Retval &= " " & dymmyNPWP.Substring(12, dymmyNPWP.Length - 12)
            Else
                Return vNPWPNo
            End If

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return vNPWPNo
        End Try
    End Function

    '2025 Coretax Tahap 2.2 tambah Tanggal akhir tax 04
    Public Function TglAkhirTax04() As Date
        Dim objSetting As New ClsTransaksiSetting
        Dim dummy As String = ""
        Try
            dummy = objSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.NPWP, ClsTransaksiSetting.EnumNamaParameter.NPWP_AkhirTax04)
            If dummy = "" Then
                objSetting.SetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.NPWP, ClsTransaksiSetting.EnumNamaParameter.NPWP_AkhirTax04, "12/31/2026")
                Return CDate("12/31/2026")
            Else
                Return CDate(dummy)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
