Option Explicit On
Option Strict Off

Imports System.Globalization
Public Class ClsMasterKaryawan_Item

    Public Enum EnumHubKeluarga As Integer
        Husband
        Wife
        Son
        Daughter
        Brother
        Sister
        Father
        Mother
        Uncle
        Aunt
        Other
    End Enum

    Public Enum EnumSEx As Integer
        Male
        Female
    End Enum

    Public Function GetIdentifierUrutan(ByVal vDGVItem As DataGridView, _
   ByVal vColIdentifier As Integer, ByVal baru As Boolean, _
   ByVal vPrimaryTransaksi As Double, ByVal vIndexCell As Integer) As String
        Dim Row As DataGridViewRow
        Try
            vDGVItem.AllowUserToAddRows = False
            Dim Identifier As String = ""
            Dim UrutanMax As Double = 0
            If baru = True Then
                If vDGVItem.Rows.Count > 0 Then
                    For i As Integer = 0 To vDGVItem.Rows.Count - 1
                        Row = vDGVItem.Rows(i)
                        If Row.Index <> vIndexCell Then
                            With Row
                                If i = 0 Then
                                    UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1))
                                Else
                                    If CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1)) > UrutanMax Then
                                        UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-", 1) + 1))
                                    End If
                                End If
                            End With
                        End If
                    Next
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + CStr(UrutanMax + 1))
                Else
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + "1")
                End If
            Else
                If vDGVItem.Rows.Count > 0 Then
                    For i As Integer = 0 To vDGVItem.Rows.Count - 1
                        Row = vDGVItem.Rows(i)
                        If Row.Index <> vIndexCell Then
                            With Row
                                If i = 0 Then
                                    UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1))
                                Else
                                    If CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1)) > UrutanMax Then
                                        UrutanMax = CDbl(.Cells(vColIdentifier).Value.ToString.Substring(CStr(.Cells(vColIdentifier).Value).IndexOf("-") + 1))
                                    End If
                                End If
                            End With
                        End If
                    Next
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + CStr(UrutanMax + 1))
                Else
                    Identifier = CStr(vPrimaryTransaksi.ToString + "-" + "1")
                End If
            End If
            vDGVItem.AllowUserToAddRows = True

            Return Identifier
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub SimpanFamily(ByVal vDgv As DataGridView, ByVal vPrimaryKaryawan As Double, ByVal frm As FrmMasterKaryawan)
        Dim Dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_itemTableAdapter
        Dim row As DataGridViewRow
        Dim IndexSex As Integer = -1
        Dim IndexHubKel As Integer = -1
        Dim tgl As Date = Now
        Dim i As Integer = 1
        Dim CboDgv As DataGridViewComboBoxCell
        Dim Edu, Phone, remark, GolDarah As String
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dim IdentifierLama As String = ""
            Dim Identifier As String = ""
            Dim counter As Integer = 0
            'set identifier
            For Each row In vDgv.Rows
                With row
                    If i <> vDgv.Rows.Count Then
                        'cek apakah data identifier perlu diubah
                        IdentifierLama = row.Cells(frm.ColIdentifier.Index).Value.ToString
                        If IdentifierLama.StartsWith(vPrimaryKaryawan.ToString) = False Then
                            'nilai start index = 1 agar tanda minus tidak diperhitungkan
                            counter = IdentifierLama.Substring(IdentifierLama.IndexOf("-", 1) + 1)
                            Identifier = vPrimaryKaryawan.ToString + "-" + counter.ToString
                            'ganti data identifier lama dengan yang baru, untuk item
                            .Cells(frm.ColIdentifier.Index).Value = Identifier
                        Else
                            Identifier = IdentifierLama
                        End If
                        CboDgv = .Cells(frm.ColHubKeluarga.Index)
                        IndexHubKel = CboDgv.Items.IndexOf(.Cells(frm.ColHubKeluarga.Index).Value)
                        CboDgv = .Cells(frm.ColSex.Index)
                        IndexSex = CboDgv.Items.IndexOf(.Cells(frm.ColSex.Index).Value)
                        My.Application.ChangeCulture("id-ID")
                        My.Application.ChangeUICulture("id-ID")
                        If IsDate(row.Cells(frm.ColBirthday.Index).Value) = False Then
                            tgl = #1/1/1950#
                        Else
                            tgl = row.Cells(frm.ColBirthday.Index).Value
                        End If
                        My.Application.ChangeCulture("en-US")
                        My.Application.ChangeUICulture("en-US")
                        If String.IsNullOrEmpty(.Cells(frm.ColEducation.Index).Value) = True Then
                            Edu = ""
                        Else
                            Edu = .Cells(frm.ColEducation.Index).Value
                        End If
                        If String.IsNullOrEmpty(.Cells(frm.ColContactNo.Index).Value) = True Then
                            Phone = ""
                        Else
                            Phone = .Cells(frm.ColContactNo.Index).Value
                        End If
                        If String.IsNullOrEmpty(.Cells(frm.ColRemarks.Index).Value) = True Then
                            remark = ""
                        Else
                            remark = .Cells(frm.ColRemarks.Index).Value
                        End If
                        CboDgv = .Cells(frm.ColGolDarah.Index)
                        GolDarah = CboDgv.Items.IndexOf(.Cells(frm.ColGolDarah.Index).Value)

                        Dta.Insert(vPrimaryKaryawan, .Cells(frm.ColIdentifier.Index).Value, .Cells(frm.ColName.Index).Value, _
                        IndexHubKel, .Cells(frm.ColIsDependent.Index).Value, Edu, _
                        IndexSex, Phone, .Cells(frm.ColBirthPlace.Index).Value, _
                        tgl, remark, GolDarah)
                        i = i + 1
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Sub

    Public Sub SimpanKompetensi(ByVal vDGV As DataGridView, ByVal vPrimaryKaryawan As Double, ByVal frm As FrmMasterKaryawan)
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim row As DataGridViewRow
        Dim tgl As Date
        Dim i As Integer = 1
        Dim remark As String = ""
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vDGV.Rows
                If i <> vDGV.Rows.Count Then
                    If row.Cells(frm.ColCompetence_NextCertificationDate.Index).Value = "-" Then
                        tgl = CDate("1/1/1950")
                    Else
                        tgl = row.Cells(frm.ColCompetence_NextCertificationDate.Index).Value
                    End If

                    If String.IsNullOrEmpty(row.Cells(frm.ColCompetence_Remarks.Index).Value) = True Then
                        remark = ""
                    Else
                        remark = row.Cells(frm.ColCompetence_Remarks.Index).Value
                    End If
                    dta.Insert(vPrimaryKaryawan, row.Cells(frm.ColCompetence_Primary.Index).Value, _
                    row.Cells(frm.ColCompetence_Value.Index).Value, remark, _
                    tgl)
                    i = i + 1
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Sub
    Public Sub SimpanAttachment(ByVal vDGV As DataGridView, ByVal vPrimaryKaryawan As Double, ByVal frm As FrmMasterKaryawan)
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_item_attachmentTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowDGV As DataGridViewRow In frm.DGVAttanchment.Rows
                If rowDGV.Cells(frm.ColAtt_Description.Index).Value <> "" Then
                    dta.Insert(rowDGV.Cells(frm.ColAtt_Description.Index).Value, vPrimaryKaryawan, rowDGV.Cells(frm.ColAtt_LocationFile.Index).Value, _
                    rowDGV.Cells(frm.ColAtt_Remarks.Index).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Sub
    Public Function SimpanKontrak(ByVal vDGV As DataGridView, ByVal vPrimaryKaryawan As Double, ByVal frm As FrmMasterKaryawan) As Boolean
        Dim dta As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_item_kontrakTableAdapter


        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            For i As Integer = 0 To frm.tblItemTipeKontrak.Rows.Count - 1
                frm.rowItemTipeKontrak = frm.tblItemTipeKontrak.Rows.Item(i)
                If frm.rowItemTipeKontrak.identifier_kontrak.ToString.StartsWith("-1") Then
                    frm.rowItemTipeKontrak.identifier_kontrak = vPrimaryKaryawan & _
                    frm.rowItemTipeKontrak.identifier_kontrak.ToString.Substring(2)
                End If
                dta.Insert(vPrimaryKaryawan, frm.rowItemTipeKontrak.primary_tipe_kontrak, _
                frm.rowItemTipeKontrak.document_no, frm.rowItemTipeKontrak._date, frm.rowItemTipeKontrak.start_date, _
                frm.rowItemTipeKontrak.end_date, frm.rowItemTipeKontrak.remarks, frm.rowItemTipeKontrak.identifier_kontrak)
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function
    Public Function SimpanAttchKontrak(ByVal vDGV As DataGridView, ByVal vPrimaryKaryawan As Double, ByVal frm As FrmMasterKaryawan) As Boolean
        Dim dta As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_item_kontrak_attachmentTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            For i As Integer = 0 To frm.tblItemAttch.Rows.Count - 1
                frm.rowItemAttch = frm.tblItemAttch.Rows.Item(i)
                If frm.rowItemAttch.identifier_kontrak.ToString.StartsWith("-1") Then
                    frm.rowItemAttch.identifier_kontrak = vPrimaryKaryawan & _
                    frm.rowItemAttch.identifier_kontrak.ToString.Substring(2)
                End If
                dta.Insert(vPrimaryKaryawan, frm.rowItemAttch.file_path, _
                frm.rowItemAttch.remarks, frm.rowItemAttch.identifier_kontrak, frm.rowItemAttch.desc)
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function
    Public Function HapusItem(ByVal vPrimaryKaryawan As Double) As Boolean
        Dim dtaFamily As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_itemTableAdapter
        Dim dtaKompetensi As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim dtaKontrak As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_item_kontrakTableAdapter
        Dim dtaAttchKontrak As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_item_kontrak_attachmentTableAdapter
        Dim dtaAttachment As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_item_attachmentTableAdapter
        Try

            'hapus data item
            dtaFamily.Connection.ConnectionString = Core_Module.GetConstring
            dtaKompetensi.Connection.ConnectionString = Core_Module.GetConstring
            dtaKontrak.Connection.ConnectionString = Core_Module.GetConstring
            dtaAttchKontrak.Connection.ConnectionString = Core_Module.GetConstring
            dtaAttachment.Connection.ConnectionString = Core_Module.GetConstring

            dtaFamily.Delete(vPrimaryKaryawan)
            dtaKompetensi.Delete(vPrimaryKaryawan)
            dtaKontrak.Delete(vPrimaryKaryawan)
            dtaAttchKontrak.Delete(vPrimaryKaryawan)
            dtaAttachment.Delete(vPrimaryKaryawan)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaFamily.Dispose()
            dtaKompetensi.Dispose()
        End Try
    End Function

    Public Function IsiDGVItemFamily(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmMasterKaryawan) As Boolean
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_itemTableAdapter
        Dim TBL As New DseMasterKaryawan.tblmaster_karyawan_itemDataTable
        Dim row As DseMasterKaryawan.tblmaster_karyawan_itemRow
        Dim rowDGV As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()

            ''isikan(tabel)
            vDGVItem.AllowUserToAddRows = False
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColIdentifier.Index).Value = .identifier
                        rowDGV.Cells(Frm.ColName.Index).Value = .nama
                        CboDgv = rowDGV.Cells(Frm.ColHubKeluarga.Index)
                        rowDGV.Cells(Frm.ColHubKeluarga.Index).Value = CboDgv.Items.Item(.relationship)
                        rowDGV.Cells(Frm.ColIsDependent.Index).Value = .Is_dependent
                        rowDGV.Cells(Frm.ColEducation.Index).Value = .education
                        rowDGV.Cells(Frm.ColRemarks.Index).Value = .remarks
                        CboDgv = rowDGV.Cells(Frm.ColSex.Index)
                        rowDGV.Cells(Frm.ColSex.Index).Value = CboDgv.Items.Item(.sex)
                        rowDGV.Cells(Frm.ColContactNo.Index).Value = .phone_no
                        rowDGV.Cells(Frm.ColBirthPlace.Index).Value = .birth_place
                        My.Application.ChangeCulture("id-ID")
                        My.Application.ChangeUICulture("id-ID")
                        rowDGV.Cells(Frm.ColBirthday.Index).Value = Format(CDate(.birthday), "dd MMM yyyy")
                        My.Application.ChangeCulture("en-US")
                        My.Application.ChangeUICulture("en-US")
                        CboDgv = rowDGV.Cells(Frm.ColGolDarah.Index)
                        rowDGV.Cells(Frm.ColGolDarah.Index).Value = CboDgv.Items.Item(.Blood_Type)
                    End With
                Next
            End If
            vDGVItem.AllowUserToAddRows = True
            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Function

    Public Function IsiDGVItemKompetensi(ByVal vDGVItem As DataGridView, ByVal vPrimaryKaryawan As Double, _
    ByVal Frm As FrmMasterKaryawan, ByVal vPrimaryJob As Double) As Boolean
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim TBL As New DseMasterKaryawan.tblmaster_karyawan_item_kompetensiDataTable
        Dim row As DseMasterKaryawan.tblmaster_karyawan_item_kompetensiRow
        Dim rowDGV As DataGridViewRow
        Dim ObjKompetensi As New ClsMasterKompetensi
        Dim ObjJob As New ClsMasterJobPosition
        Dim ObjCekData As New ClsCekData
        Dim DataType As Integer = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()
            vDGVItem.AllowUserToAddRows = False
            ''isikan(tabel)
            If DTA.FillBy(TBL, vPrimaryKaryawan) > 0 Then
                For Each row In TBL.Rows
                    With row
                        DataType = ObjKompetensi.GetDataType(.primary_kompetensi)
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColCompetence_Code.Index).Value = ObjKompetensi.GetDataCode(.primary_kompetensi)
                        rowDGV.Cells(Frm.ColCempetence_Description.Index).Value = ObjKompetensi.GetDesc(.primary_kompetensi)
                        If ObjJob.GetJmlItem(vPrimaryJob, .primary_kompetensi) > 0 Then
                            rowDGV.Cells(Frm.ColCompetence_StandarValue.Index).Value = _
                            ObjJob.GetValue_CustomFormatted(vPrimaryJob, .primary_kompetensi, ObjCekData, DataType)
                        Else
                            rowDGV.Cells(Frm.ColCompetence_StandarValue.Index).Value = _
                            ObjKompetensi.GetValue_CustomFormatted(.primary_kompetensi, ObjCekData, DataType)
                        End If
                        rowDGV.Cells(Frm.ColCompetence_Value.Index).Value = .value
                        'My.Application.ChangeCulture("id-ID")
                        'My.Application.ChangeUICulture("id-ID")
                        If .next_certification_date = "1/1/1950" Then
                            rowDGV.Cells(Frm.ColCompetence_NextCertificationDate.Index).Value = "-"
                        Else
                            rowDGV.Cells(Frm.ColCompetence_NextCertificationDate.Index).Value = Format(.next_certification_date, "dd MMM yyyy")
                        End If
                        'My.Application.ChangeCulture("en-US")
                        'My.Application.ChangeUICulture("en-US")
                        rowDGV.Cells(Frm.ColCompetence_Remarks.Index).Value = .remarks
                        rowDGV.Cells(Frm.ColCompetence_Primary.Index).Value = .primary_kompetensi
                        rowDGV.Cells(Frm.ColCempetence_DataType.Index).Value = DataType
                    End With
                Next
            End If
            'Frm.UpdateStandarValue()
            vDGVItem.AllowUserToAddRows = True
            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCekData = Nothing
            ObjJob = Nothing
            ObjKompetensi = Nothing
        End Try
    End Function

    Public Function IsiDGVItemKontrak(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal Frm As FrmMasterKaryawan) As Boolean
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_item_kontrakTableAdapter
        Dim TBL As New DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrakDataTable
        Dim row As DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrakRow
        Dim rowDGV As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell
        Dim objKaryawanType As New ClsMasterKaryawanTipe

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()
            Frm.tblItemTipeKontrak.Clear()
            ''isikan(tabel)
            vDGVItem.AllowUserToAddRows = False
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        CboDgv = rowDGV.Cells(Frm.ColContract_Type.Index)
                        CboDgv.DataSource = Frm.TBLKontrak.DefaultView
                        CboDgv.DisplayMember = "Code"
                        CboDgv.ValueMember = "Primary"
                        rowDGV.Cells(Frm.ColContract_Type.Index).Value = .primary_tipe_kontrak
                        rowDGV.Cells(Frm.ColContract_DocNo.Index).Value = .document_no
                        rowDGV.Cells(Frm.ColContract_Date.Index).Value = ._date 'Format(CDate(._date), "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColContract_Start.Index).Value = .start_date 'Format(CDate(.start_date), "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColContract_End.Index).Value = .end_date 'Format(CDate(.end_date), "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColContract_Remarks.Index).Value = .remarks
                        rowDGV.Cells(Frm.ColContract_PrimaryType.Index).Value = .primary_tipe_kontrak
                        rowDGV.Cells(Frm.ColContract_IdentifierContract.Index).Value = .identifier_kontrak
                        objKaryawanType.TambahBarisDGVKontrak(Frm.tblItemTipeKontrak, .primary_tipe_kontrak, _
                        vPrimaryTransaksi, .document_no, ._date, .start_date, .end_date, .remarks, .identifier_kontrak, Frm)
                    End With
                Next
            End If
            vDGVItem.AllowUserToAddRows = True
            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Function
    Public Function IsiDGVItemAttchKontrak(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
   ByVal Frm As FrmMasterKaryawan) As Boolean
        Dim DTA As New DseMasterKaryawanTipeTableAdapters.tblmaster_karyawan_item_kontrak_attachmentTableAdapter
        Dim TBL As New DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrak_attachmentDataTable
        Dim row As DseMasterKaryawanTipe.tblmaster_karyawan_item_kontrak_attachmentRow
        Dim rowDGV As DataGridViewRow
        Dim objKaryawanType As New ClsMasterKaryawanTipe

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()
            Frm.tblItemAttch.Clear()
            ''isikan(tabel)
            vDGVItem.AllowUserToAddRows = False
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColKontrakAtt_Description.Index).Value = .desc
                        rowDGV.Cells(Frm.ColKontrakAtt_LocationFile.Index).Value = .file_path
                        rowDGV.Cells(Frm.ColKontrakAtt_Preview.Index).Value = "preview"
                        rowDGV.Cells(Frm.ColKontrakAtt_BtnOpenDelete.Index).Value = "remove"
                        rowDGV.Cells(Frm.ColKontrakAtt_Remarks.Index).Value = .remarks
                        objKaryawanType.TambahBarisDGVAttchKontrak(Frm.tblItemAttch, vPrimaryTransaksi, _
                        .desc, .file_path, .remarks, .identifier_kontrak, Frm)
                    End With
                Next
                Frm.reloadDataTableKontrak(Frm.DGVKontrak.Item(Frm.ColContract_IdentifierContract.Index, 0).Value)
            End If
            vDGVItem.AllowUserToAddRows = True
            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Function
    Public Function IsiDGVItemAttachment(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
   ByVal Frm As FrmMasterKaryawan) As Boolean
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_item_attachmentTableAdapter
        Dim TBL As New DseMasterKaryawan.tblmaster_karyawan_item_attachmentDataTable
        Dim row As DseMasterKaryawan.tblmaster_karyawan_item_attachmentRow
        Dim rowDGV As DataGridViewRow


        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.Rows.Clear()
            ''isikan(tabel)
            vDGVItem.AllowUserToAddRows = False
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColAtt_Description.Index).Value = .desc
                        rowDGV.Cells(Frm.ColAtt_LocationFile.Index).Value = .file_path
                        rowDGV.Cells(Frm.ColAtt_Preview.Index).Value = "preview"
                        rowDGV.Cells(Frm.ColAtt_BtnOpenDelete.Index).Value = "remove"
                        rowDGV.Cells(Frm.ColAtt_Remarks.Index).Value = .remarks
                    End With
                Next
            End If
            vDGVItem.AllowUserToAddRows = True
            'set nomer
            ''ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Function
End Class
