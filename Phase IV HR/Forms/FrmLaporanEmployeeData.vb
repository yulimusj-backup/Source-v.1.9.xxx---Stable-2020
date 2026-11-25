Option Explicit On
Option Strict Off

Imports Core_Module, System.Windows.Forms, System.Drawing
Public Class FrmLaporanEmployeeData
    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Dim TBLJobPosition As New DseLaporanEmployeeData.tblmaster_jobpositionDataTable
    Dim TBLDepartement As New DseLaporanEmployeeData.tblmaster_departemenDataTable
    Dim TBLEmployee As New DseLaporanEmployeeData.tblmaster_karyawanDataTable
    Dim RowJobPosition As DseLaporanEmployeeData.tblmaster_jobpositionRow
    Dim RowDepartement As DseLaporanEmployeeData.tblmaster_departemenRow
    Dim RowEmployee As DseLaporanEmployeeData.tblmaster_karyawanRow
    Dim ObjEmp As New Core_Module.ClsMasterKaryawan
    Dim ObjDept As New Core_Module.ClsMasterDepartement
    Dim ObjJobPost As New Core_Module.ClsMasterJobPosition
    Dim IndexEmp As Integer = 0
    Dim IndexDept As Integer = 0
    Dim IndexJobPost As Integer = 0
    Dim ListEmp As New List(Of Double)
    Dim ListJob As New List(Of Double)
    Dim ListDept As New List(Of Double)

    Dim Ketemu As Boolean = False
    Dim Buffer As String = ""
    Private Enum EnumTipe As Integer
        JobPosition
        Departement
    End Enum

    Private Sub IsiListBoxDepartement(ByVal lbo As Windows.Forms.CheckedListBox, ByVal vTBLDepartement As DseLaporanEmployeeData.tblmaster_departemenDataTable)
        Dim DTADepartement As New DseLaporanEmployeeDataTableAdapters.tblmaster_departemenTableAdapter
        Try
            DTADepartement.Connection.ConnectionString = Core_Module.GetConstring
            vTBLDepartement.Rows.Clear()
            DTADepartement.Fill(vTBLDepartement)
            lbo.DataSource = Nothing
            lbo.DataSource = vTBLDepartement.DefaultView

            'isikan No Event 
            lbo.DisplayMember = "Departement"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTADepartement.Dispose()
        End Try
    End Sub
   

    Private Sub IsiListBoxJobPosition(ByVal lbo As Windows.Forms.CheckedListBox, ByVal vTBLJobPosition As DseLaporanEmployeeData.tblmaster_jobpositionDataTable)
        Dim DTAJobPosition As New DseLaporanEmployeeDataTableAdapters.tblmaster_jobpositionTableAdapter
        Try
            DTAJobPosition.Connection.ConnectionString = Core_Module.GetConstring
            vTBLJobPosition.Rows.Clear()
            DTAJobPosition.Fill(vTBLJobPosition)
            lbo.DataSource = Nothing
            lbo.DataSource = vTBLJobPosition.DefaultView

            'isikan No Event 
            lbo.DisplayMember = "Position"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAJobPosition.Dispose()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KendalikanKontrol(ByVal vEnum As EnumTipe, ByVal vAktif As Boolean)
        Try
            If vEnum = EnumTipe.JobPosition Then
                If vAktif = True Then
                    For i As Integer = 0 To Me.clbJobPosition.Items.Count - 1
                        Me.clbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
                    Next
                    Me.Lbl_JB_Checked.Enabled = True
                    Me.Lbl_JB_Unchecked.Enabled = True
                    Me.clbJobPosition.Enabled = True
                Else
                    For i As Integer = 0 To Me.clbJobPosition.Items.Count - 1
                        Me.clbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
                    Next
                    Me.Lbl_JB_Checked.Enabled = False
                    Me.Lbl_JB_Unchecked.Enabled = False
                    Me.clbJobPosition.Enabled = False
                End If
            End If
            If vEnum = EnumTipe.Departement Then
                If vAktif = True Then
                    For i As Integer = 0 To Me.ClbDepartement.Items.Count - 1
                        Me.ClbDepartement.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
                    Next
                    Me.Lbl_Dept_Checked.Enabled = True
                    Me.Lbl_Dept_Unchecked.Enabled = True
                    Me.ClbDepartement.Enabled = True
                Else
                    For i As Integer = 0 To Me.ClbDepartement.Items.Count - 1
                        Me.ClbDepartement.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
                    Next
                    Me.Lbl_Dept_Checked.Enabled = False
                    Me.Lbl_Dept_Unchecked.Enabled = False
                    Me.ClbDepartement.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ListJobPosition(ByVal vListJP As List(Of Double), ByVal vExecute As Boolean)
        Dim RowView As Data.DataRowView
        Dim PrimaryJobPosition As Double
        Try
            vListJP.Clear()
            If vExecute = True Then
                For Each RowView In Me.clbJobPosition.CheckedItems
                    PrimaryJobPosition = RowView.Row.Item("primary")
                    If vListJP.Contains(PrimaryJobPosition) = False Then
                        vListJP.Add(PrimaryJobPosition)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub ListJobPositionBaru()
        Try
            Me.ListJob.Clear()
            If Me.dgvJobPosition.Rows.Count - 1 <> 0 Then
                For i As Integer = 0 To Me.dgvJobPosition.Rows.Count - 2
                    If Me.ListJob.Contains(Me.dgvJobPosition.Item(Me.ColJobPosisitonPrimary.Index, i).Value) = False Then
                        Me.ListJob.Add(Me.dgvJobPosition.Item(Me.ColJobPosisitonPrimary.Index, i).Value)
                    End If
                Next
            Else
                For x As Integer = 0 To Me.TBLJobPosition.Count - 1
                    RowJobPosition = TBLJobPosition.Rows(x)
                    If Me.ListJob.Contains(RowJobPosition.primary) = False Then
                        Me.ListJob.Add(RowJobPosition.primary)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub ListEmployeeBaru()

        Try
            Me.ListEmp.Clear()
            If Me.dgvEmployee.Rows.Count - 1 <> 0 Then
                For i As Integer = 0 To Me.dgvEmployee.Rows.Count - 2
                    If Me.ListEmp.Contains(Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, i).Value) = False Then
                        Me.ListEmp.Add(Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, i).Value)
                    End If
                Next
            Else
                For x As Integer = 0 To Me.TBLEmployee.Count - 1
                    RowEmployee = TBLEmployee.Rows(x)
                    If Me.ListEmp.Contains(RowEmployee.Primary) = False Then
                        Me.ListEmp.Add(RowEmployee.Primary)
                    End If
                Next

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub ListDeptBaru()
        Try
            Me.ListDept.Clear()
            If Me.dgvDepartment.Rows.Count - 1 <> 0 Then
                For i As Integer = 0 To Me.dgvDepartment.Rows.Count - 2
                    If Me.ListDept.Contains(Me.dgvDepartment.Item(Me.ColDeptPrimary.Index, i).Value) = False Then
                        Me.ListDept.Add(Me.dgvDepartment.Item(Me.ColDeptPrimary.Index, i).Value)
                    End If
                Next
            Else
                For x As Integer = 0 To Me.TBLDepartement.Count - 1
                    RowDepartement = TBLDepartement.Rows(x)
                    If Me.ListDept.Contains(RowDepartement.Primary) = False Then
                        Me.ListDept.Add(RowDepartement.Primary)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub ListDepartement(ByVal vListDept As List(Of Double), ByVal vExecute As Boolean)
        Dim RowView As Data.DataRowView
        Dim PrimaryDepartement As Double
        Try
            vListDept.Clear()
            If vExecute = True Then
                For Each RowView In Me.ClbDepartement.CheckedItems
                    PrimaryDepartement = RowView.Row.Item("Primary")
                    If vListDept.Contains(PrimaryDepartement) = False Then
                        vListDept.Add(PrimaryDepartement)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatLaporan(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vChkJobPosition As Boolean, _
    ByVal vChkDepartement As Boolean, ByVal vApakahDataBenar As Boolean)
        Dim Report1 As New CrpLaporanEmployee_Format1
        Dim Report2 As New CrpLaporanEmployee_Format2
        Dim Report3 As New CrpLaporanEmployee_FamilyData
        Dim DS As New DseLaporanEmployeeData
        Dim DTA_Format1 As New DseLaporanEmployeeDataTableAdapters.Format_1TableAdapter
        Dim TBL_Format1 As DseLaporanEmployeeData.Format_1DataTable = DS.Format_1
        Dim Row_Format1 As DseLaporanEmployeeData.Format_1Row
        Dim DTA_Format2 As New DseLaporanEmployeeDataTableAdapters.Format_2TableAdapter
        Dim TBL_Format2 As DseLaporanEmployeeData.Format_2DataTable = DS.Format_2
        Dim Row_Format2 As DseLaporanEmployeeData.Format_2Row
        Dim DTA_Family As New DseLaporanEmployeeDataTableAdapters.Anggota_KeluargaTableAdapter
        Dim TBL_Family As DseLaporanEmployeeData.Anggota_KeluargaDataTable = DS.Anggota_Keluarga
        Dim Row_Family As DseLaporanEmployeeData.Anggota_KeluargaRow
        'Dim Row_Family_Minus_Satu As DseLaporanEmployeeData.Anggota_KeluargaRow
        Dim TBL_Dummy_Format1 As New DseLaporanEmployeeData.Format_1DataTable
        Dim Row_Dummy_Format1 As DseLaporanEmployeeData.Format_1Row
        Dim TBL_Dummy_Format2 As New DseLaporanEmployeeData.Format_2DataTable
        Dim Row_Dummy_Format2 As DseLaporanEmployeeData.Format_2Row
        Dim TBL_Dummy_Family As New DseLaporanEmployeeData.Anggota_KeluargaDataTable
        Dim Row_Dummy_Family As DseLaporanEmployeeData.Anggota_KeluargaRow
        Dim mListJobPosition As New List(Of Double)
        Dim mListDepartement As New List(Of Double)
        Dim mListPrimary As New List(Of Double)

        Dim mListIdentifier As New List(Of String)
        Dim mListGroup As New List(Of String)
        Dim mListGroupDetail As New List(Of String)

        Dim listPrimaryDummy As New List(Of String)
        Dim DV As Data.DataRowView
        Dim ExecuteJobPostion As Boolean = Me.ChkJobPosition.Checked
        Dim ExecuteDepartement As Boolean = Me.ChkDepartement.Checked
        Dim PrimaryJobPosition As Double
        Dim PrimaryDepartement As Double
        Dim PrimaryEmployee As Double
        Dim Filter As String
        Dim i As Integer = 0
        Dim ix As Integer = 0
        Dim tampil As Boolean = False
        Try
            DTA_Family.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Format1.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Format2.Connection.ConnectionString = Core_Module.GetConstring
            TBL_Format1.Rows.Clear()
            TBL_Format2.Rows.Clear()
            TBL_Family.Rows.Clear()
            TBL_Dummy_Format1.Rows.Clear()
            TBL_Dummy_Format2.Rows.Clear()
            TBL_Dummy_Family.Rows.Clear()
            TBL_Dummy_Format1.DefaultView.RowFilter = ""
            TBL_Dummy_Format1.DefaultView.Sort = ""
            TBL_Dummy_Format2.DefaultView.RowFilter = ""
            TBL_Dummy_Format2.DefaultView.Sort = ""
            TBL_Family.DefaultView.RowFilter = ""
            TBL_Family.DefaultView.Sort = ""
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
            If vApakahDataBenar = True Then
                'Buat List data yang di cek list
                'Me.ListJobPosition(mListJobPosition, ExecuteJobPostion)
                'Me.ListDepartement(mListDepartement, ExecuteDepartement)
                Me.ListEmployeeBaru()
                Me.ListDeptBaru()
                Me.ListJobPositionBaru()
                'Fill all data berdasarkan format-nya
                'Format 1
                'hilangkan data yg sama, jadikan semua primary unik tidak ada yg sama
                listPrimaryDummy.Clear()
                If Me.RboFormat1.Checked = True Then
                    DTA_Format1.Fill(TBL_Format1)
                    While ix <= TBL_Format1.Rows.Count - 1
                        Row_Format1 = TBL_Format1.Rows(ix)
                        If listPrimaryDummy.Contains(Row_Format1.Primary) = False Then
                            listPrimaryDummy.Add(Row_Format1.Primary)
                            ix += 1
                        Else
                            TBL_Format1.Rows.Remove(Row_Format1)
                        End If
                    End While
                End If
                'Format 2
                If Me.RboFormat2.Checked = True Then
                    DTA_Format2.Fill(TBL_Format2)
                    While ix <= TBL_Format2.Rows.Count - 1
                        Row_Format2 = TBL_Format2.Rows(ix)
                        If listPrimaryDummy.Contains(Row_Format2.Primary) = False Then
                            listPrimaryDummy.Add(Row_Format2.Primary)
                            ix += 1
                        Else
                            TBL_Format2.Rows.Remove(Row_Format2)
                        End If
                    End While
                End If
                'Family data
                If Me.RboFamilyData.Checked = True Then
                    DTA_Family.Fill(TBL_Family)
                    While ix <= TBL_Family.Rows.Count - 1
                        Row_Family = TBL_Family.Rows(ix)
                        If listPrimaryDummy.Contains(Row_Family.Primary.ToString + "," + Row_Family.Nama_Anggota_Keluarga) = False Then
                            listPrimaryDummy.Add(Row_Family.Primary.ToString + "," + Row_Family.Nama_Anggota_Keluarga)
                            ix += 1
                        Else
                            TBL_Family.Rows.Remove(Row_Family)
                        End If
                    End While
                End If
            
              
                
               
                'Filter data menggunakan data view :
                'a. Filter berdasarkan check box sesuaikan dengan kondisi (joblist saja, departement saja, atau keduanya)
                If Me.RboFormat1.Checked = True Then
                    '@ tambahan filteran active or not active for employee data
                    Me.FilterAktifEmployee(TBL_Format1)
                    If vChkJobPosition = True And vChkDepartement = True Then
                        For Each PrimaryJobPosition In Me.ListJob
                            TBL_Format1.DefaultView.RowFilter = ""
                            Filter = "[primary position] = " + PrimaryJobPosition.ToString + ""
                            TBL_Format1.DefaultView.RowFilter = Filter
                            If TBL_Format1.DefaultView.Count > 0 Then
                                For Each DV In TBL_Format1.DefaultView
                                    Row_Format1 = DV.Row
                                    If Me.cboPeriod.Text = "All Data" Then
                                        tampil = True
                                    Else
                                        If Row_Format1.DOC_NO <> "" Then
                                            If Me.cboPeriod.Text = "Contract Date" Then
                                                If Me.dtpFrom.Value.Date <= Row_Format1._date And Me.dtpTo.Value.Date >= Row_Format1._date Then
                                                    tampil = True
                                                Else
                                                    tampil = False
                                                End If
                                            ElseIf Me.cboPeriod.Text = "Contract End Date" Then
                                                If Me.dtpFrom.Value.Date <= Row_Format1.start_date And Me.dtpTo.Value.Date >= Row_Format1.start_date Then
                                                    tampil = True
                                                Else
                                                    tampil = False
                                                End If
                                            ElseIf Me.cboPeriod.Text = "Contract Start Date" Then
                                                If Me.dtpFrom.Value.Date <= Row_Format1.end_date And Me.dtpTo.Value.Date >= Row_Format1.end_date Then
                                                    tampil = True
                                                Else
                                                    tampil = False
                                                End If
                                            End If
                                        Else
                                            tampil = False
                                        End If
                                    End If
                                    If tampil = True Then
                                        If mListPrimary.Contains(Row_Format1.Primary) = False Then
                                            mListPrimary.Add(Row_Format1.Primary)
                                            Row_Dummy_Format1 = TBL_Dummy_Format1.NewFormat_1Row
                                            For a As Integer = 0 To TBL_Dummy_Format1.Columns.Count - 1
                                                Row_Dummy_Format1.Item(a) = Row_Format1.Item(a)
                                            Next
                                            TBL_Dummy_Format1.Rows.Add(Row_Dummy_Format1)
                                        End If
                                    End If
                                Next
                            End If
                        Next
                        Filter = ""
                        For Each PrimaryDepartement In Me.ListDept
                            If Filter = "" Then
                                Filter = "[primary departemen] <> " + PrimaryDepartement.ToString + ""
                            Else
                                Filter += " AND [primary departemen] <> " + PrimaryDepartement.ToString + ""
                            End If
                        Next
                        TBL_Dummy_Format1.DefaultView.RowFilter = Filter
                        If TBL_Dummy_Format1.DefaultView.Count > 0 Then
                            While i <= TBL_Dummy_Format1.DefaultView.Count - 1
                                Row_Dummy_Format1 = TBL_Dummy_Format1.DefaultView(i).Row
                                TBL_Dummy_Format1.Rows.Remove(Row_Dummy_Format1)
                            End While
                        End If
                        Filter = ""
                        i = 0
                        For Each PrimaryEmployee In Me.ListEmp
                            If Filter = "" Then
                                Filter = "[primary] <> " + PrimaryEmployee.ToString + ""
                            Else
                                Filter += " AND [primary] <> " + PrimaryEmployee.ToString + ""
                            End If
                        Next
                        TBL_Dummy_Format1.DefaultView.RowFilter = Filter
                        If TBL_Dummy_Format1.DefaultView.Count > 0 Then
                            While i <= TBL_Dummy_Format1.DefaultView.Count - 1
                                Row_Dummy_Format1 = TBL_Dummy_Format1.DefaultView(i).Row
                                TBL_Dummy_Format1.Rows.Remove(Row_Dummy_Format1)
                            End While
                        End If

                        'Sorting data
                        TBL_Format1.DefaultView.RowFilter = ""
                        TBL_Format1.Rows.Clear()
                        TBL_Dummy_Format1.DefaultView.RowFilter = ""
                        TBL_Dummy_Format1.DefaultView.Sort = "[Dept Desc] ASC, [Nama Karyawan] ASC"
                        For Each DV In TBL_Dummy_Format1.DefaultView
                            Row_Dummy_Format1 = DV.Row
                            Row_Format1 = TBL_Format1.NewFormat_1Row
                            For a As Integer = 0 To TBL_Format1.Columns.Count - 1
                                If a = 7 Then
                                    Row_Format1.Item(a) = Row_Dummy_Format1.Item(a).ToString.ToLower
                                Else
                                    Row_Format1.Item(a) = Row_Dummy_Format1.Item(a)
                                End If
                            Next
                            TBL_Format1.Rows.Add(Row_Format1)
                        Next
                        'Else
                        '    If vChkJobPosition = True Then
                        '        For Each PrimaryJobPosition In mListJobPosition
                        '            TBL_Format1.DefaultView.RowFilter = ""
                        '            Filter = "[primary position] = " + PrimaryJobPosition.ToString + ""
                        '            TBL_Format1.DefaultView.RowFilter = Filter
                        '            If TBL_Format1.DefaultView.Count > 0 Then
                        '                For Each DV In TBL_Format1.DefaultView
                        '                    Row_Format1 = DV.Row
                        '                    If mListPrimary.Contains(Row_Format1.Primary) = False Then
                        '                        mListPrimary.Add(Row_Format1.Primary)
                        '                        Row_Dummy_Format1 = TBL_Dummy_Format1.NewFormat_1Row
                        '                        For a As Integer = 0 To TBL_Dummy_Format1.Columns.Count - 1
                        '                            Row_Dummy_Format1.Item(a) = Row_Format1.Item(a)
                        '                        Next
                        '                        TBL_Dummy_Format1.Rows.Add(Row_Dummy_Format1)
                        '                    End If
                        '                Next
                        '            End If
                        '        Next
                        '        TBL_Format1.DefaultView.RowFilter = ""
                        '        TBL_Format1.Rows.Clear()
                        '        TBL_Dummy_Format1.DefaultView.RowFilter = ""
                        '        TBL_Dummy_Format1.DefaultView.Sort = "[Dept Desc] ASC, [Nama Karyawan] ASC"
                        '        For Each DV In TBL_Dummy_Format1.DefaultView
                        '            Row_Dummy_Format1 = DV.Row
                        '            Row_Format1 = TBL_Format1.NewFormat_1Row
                        '            For a As Integer = 0 To TBL_Format1.Columns.Count - 1
                        '                Row_Format1.Item(a) = Row_Dummy_Format1.Item(a)
                        '            Next
                        '            TBL_Format1.Rows.Add(Row_Format1)
                        '        Next
                        '    ElseIf vChkDepartement = True Then
                        '        For Each PrimaryDepartement In mListDepartement
                        '            TBL_Format1.DefaultView.RowFilter = ""
                        '            Filter = "[primary departemen] = " + PrimaryDepartement.ToString + ""
                        '            TBL_Format1.DefaultView.RowFilter = Filter
                        '            If TBL_Format1.DefaultView.Count > 0 Then
                        '                For Each DV In TBL_Format1.DefaultView
                        '                    Row_Format1 = DV.Row
                        '                    If mListPrimary.Contains(Row_Format1.Primary) = False Then
                        '                        mListPrimary.Add(Row_Format1.Primary)
                        '                        Row_Dummy_Format1 = TBL_Dummy_Format1.NewFormat_1Row
                        '                        For a As Integer = 0 To TBL_Dummy_Format1.Columns.Count - 1
                        '                            Row_Dummy_Format1.Item(a) = Row_Format1.Item(a)
                        '                        Next
                        '                        TBL_Dummy_Format1.Rows.Add(Row_Dummy_Format1)
                        '                    End If
                        '                Next
                        '            End If
                        '        Next
                        '        TBL_Format1.DefaultView.RowFilter = ""
                        '        TBL_Format1.Rows.Clear()
                        '        TBL_Dummy_Format1.DefaultView.RowFilter = ""
                        '        TBL_Dummy_Format1.DefaultView.Sort = "[Dept Desc] ASC, [Nama Karyawan] ASC"
                        '        For Each DV In TBL_Dummy_Format1.DefaultView
                        '            Row_Dummy_Format1 = DV.Row
                        '            Row_Format1 = TBL_Format1.NewFormat_1Row
                        '            For a As Integer = 0 To TBL_Format1.Columns.Count - 1
                        '                If a = 7 Then
                        '                    Row_Format1.Item(a) = Row_Dummy_Format1.Item(a).ToString.ToLower
                        '                Else
                        '                    Row_Format1.Item(a) = Row_Dummy_Format1.Item(a)
                        '                End If
                        '            Next
                        '            TBL_Format1.Rows.Add(Row_Format1)
                        '        Next
                        '    End If
                    End If
                End If
                If Me.RboFormat2.Checked = True Then
                    '@ tambahan filteran active or not active for employee data
                    Me.FilterAktifEmployee(TBL_Format2)
                    If vChkJobPosition = True And vChkDepartement = True Then
                        For Each PrimaryJobPosition In Me.ListJob
                            TBL_Format2.DefaultView.RowFilter = ""
                            Filter = "[primary position] = " + PrimaryJobPosition.ToString + ""
                            TBL_Format2.DefaultView.RowFilter = Filter
                            If TBL_Format2.DefaultView.Count > 0 Then
                                For Each DV In TBL_Format2.DefaultView
                                    Row_Format2 = DV.Row
                                    If Me.cboPeriod.Text = "All Data" Then
                                        tampil = True
                                    Else
                                        If Row_Format2.DOC_No <> "" Then
                                            If Me.cboPeriod.Text = "Contract Date" Then
                                                If Me.dtpFrom.Value.Date <= Row_Format2._date And Me.dtpTo.Value.Date >= Row_Format2._date Then
                                                    tampil = True
                                                Else
                                                    tampil = False
                                                End If
                                            ElseIf Me.cboPeriod.Text = "Contract End Date" Then
                                                If Me.dtpFrom.Value.Date <= Row_Format2.start_date And Me.dtpTo.Value.Date >= Row_Format2.start_date Then
                                                    tampil = True
                                                Else
                                                    tampil = False
                                                End If
                                            ElseIf Me.cboPeriod.Text = "Contract Start Date" Then
                                                If Me.dtpFrom.Value.Date <= Row_Format2.end_date And Me.dtpTo.Value.Date >= Row_Format2.end_date Then
                                                    tampil = True
                                                Else
                                                    tampil = False
                                                End If
                                            End If
                                        Else
                                            tampil = False
                                        End If
                                    End If
                                    If tampil = True Then
                                        If mListPrimary.Contains(Row_Format2.Primary) = False Then
                                            mListPrimary.Add(Row_Format2.Primary)
                                            Row_Dummy_Format2 = TBL_Dummy_Format2.NewFormat_2Row
                                            For a As Integer = 0 To TBL_Dummy_Format2.Columns.Count - 1
                                                Row_Dummy_Format2.Item(a) = Row_Format2.Item(a)
                                            Next
                                            TBL_Dummy_Format2.Rows.Add(Row_Dummy_Format2)
                                        End If
                                    End If
                                Next
                            End If
                        Next
                        Filter = ""
                        For Each PrimaryDepartement In Me.ListDept
                            If Filter = "" Then
                                Filter = "[primary departemen] <> " + PrimaryDepartement.ToString + ""
                            Else
                                Filter += " AND [primary departemen] <> " + PrimaryDepartement.ToString + ""
                            End If
                        Next
                        TBL_Dummy_Format2.DefaultView.RowFilter = Filter
                        If TBL_Dummy_Format2.DefaultView.Count > 0 Then
                            While i <= TBL_Dummy_Format2.DefaultView.Count - 1
                                Row_Dummy_Format2 = TBL_Dummy_Format2.DefaultView(i).Row
                                TBL_Dummy_Format2.Rows.Remove(Row_Dummy_Format2)
                            End While
                        End If
                        Filter = ""
                        i = 0
                        For Each PrimaryEmployee In Me.ListEmp
                            If Filter = "" Then
                                Filter = "[primary] <> " + PrimaryEmployee.ToString + ""
                            Else
                                Filter += " AND [primary] <> " + PrimaryEmployee.ToString + ""
                            End If
                        Next
                        TBL_Dummy_Format2.DefaultView.RowFilter = Filter
                        If TBL_Dummy_Format2.DefaultView.Count > 0 Then
                            While i <= TBL_Dummy_Format2.DefaultView.Count - 1
                                Row_Dummy_Format2 = TBL_Dummy_Format2.DefaultView(i).Row
                                TBL_Dummy_Format2.Rows.Remove(Row_Dummy_Format2)
                            End While
                        End If
                        'Sorting data
                        TBL_Format2.DefaultView.RowFilter = ""
                        TBL_Format2.Rows.Clear()
                        TBL_Dummy_Format2.DefaultView.RowFilter = ""
                        TBL_Dummy_Format2.DefaultView.Sort = "[Dept Desc] ASC, [Nama Karyawan] ASC"
                        For Each DV In TBL_Dummy_Format2.DefaultView
                            Row_Dummy_Format2 = DV.Row
                            Row_Format2 = TBL_Format2.NewFormat_2Row
                            For a As Integer = 0 To TBL_Format2.Columns.Count - 1
                                Row_Format2.Item(a) = Row_Dummy_Format2.Item(a)
                            Next
                            TBL_Format2.Rows.Add(Row_Format2)
                        Next
                        'Else
                        '    If vChkJobPosition = True Then
                        '        For Each PrimaryJobPosition In mListJobPosition
                        '            TBL_Format2.DefaultView.RowFilter = ""
                        '            Filter = "[primary position] = " + PrimaryJobPosition.ToString + ""
                        '            TBL_Format2.DefaultView.RowFilter = Filter
                        '            If TBL_Format2.DefaultView.Count > 0 Then
                        '                For Each DV In TBL_Format2.DefaultView
                        '                    Row_Format2 = DV.Row
                        '                    If mListPrimary.Contains(Row_Format2.Primary) = False Then
                        '                        mListPrimary.Add(Row_Format2.Primary)
                        '                        Row_Dummy_Format2 = TBL_Dummy_Format2.NewFormat_2Row
                        '                        For a As Integer = 0 To TBL_Dummy_Format2.Columns.Count - 1
                        '                            Row_Dummy_Format2.Item(a) = Row_Format2.Item(a)
                        '                        Next
                        '                        TBL_Dummy_Format2.Rows.Add(Row_Dummy_Format2)
                        '                    End If
                        '                Next
                        '            End If
                        '        Next
                        '        TBL_Format2.DefaultView.RowFilter = ""
                        '        TBL_Format2.Rows.Clear()
                        '        TBL_Dummy_Format2.DefaultView.RowFilter = ""
                        '        TBL_Dummy_Format2.DefaultView.Sort = "[Dept Desc] ASC, [Nama Karyawan] ASC"
                        '        For Each DV In TBL_Dummy_Format2.DefaultView
                        '            Row_Dummy_Format2 = DV.Row
                        '            Row_Format2 = TBL_Format2.NewFormat_2Row
                        '            For a As Integer = 0 To TBL_Format2.Columns.Count - 1
                        '                Row_Format2.Item(a) = Row_Dummy_Format2.Item(a)
                        '            Next
                        '            TBL_Format2.Rows.Add(Row_Format2)
                        '        Next
                        '    ElseIf vChkDepartement = True Then
                        '        For Each PrimaryDepartement In mListDepartement
                        '            TBL_Format2.DefaultView.RowFilter = ""
                        '            Filter = "[primary departemen] = " + PrimaryDepartement.ToString + ""
                        '            TBL_Format2.DefaultView.RowFilter = Filter
                        '            If TBL_Format2.DefaultView.Count > 0 Then
                        '                For Each DV In TBL_Format2.DefaultView
                        '                    Row_Format2 = DV.Row
                        '                    If mListPrimary.Contains(Row_Format2.Primary) = False Then
                        '                        mListPrimary.Add(Row_Format2.Primary)
                        '                        Row_Dummy_Format2 = TBL_Dummy_Format2.NewFormat_2Row
                        '                        For a As Integer = 0 To TBL_Dummy_Format2.Columns.Count - 1
                        '                            Row_Dummy_Format2.Item(a) = Row_Format2.Item(a)
                        '                        Next
                        '                        TBL_Dummy_Format2.Rows.Add(Row_Dummy_Format2)
                        '                    End If
                        '                Next
                        '            End If
                        '        Next
                        '        TBL_Format2.DefaultView.RowFilter = ""
                        '        TBL_Format2.Rows.Clear()
                        '        TBL_Dummy_Format2.DefaultView.RowFilter = ""
                        '        TBL_Dummy_Format2.DefaultView.Sort = "[Dept Desc] ASC, [Nama Karyawan] ASC"
                        '        For Each DV In TBL_Dummy_Format2.DefaultView
                        '            Row_Dummy_Format2 = DV.Row
                        '            Row_Format2 = TBL_Format2.NewFormat_2Row
                        '            For a As Integer = 0 To TBL_Format2.Columns.Count - 1
                        '                Row_Format2.Item(a) = Row_Dummy_Format2.Item(a)
                        '            Next
                        '            TBL_Format2.Rows.Add(Row_Format2)
                        '        Next
                        '    End If
                    End If
                End If
                If Me.RboFamilyData.Checked = True Then
                    Me.FilterAktifEmployee(TBL_Family)
                    If vChkJobPosition = True And vChkDepartement = True Then
                        For Each PrimaryJobPosition In Me.ListJob
                            TBL_Family.DefaultView.RowFilter = ""
                            Filter = "[primary position] = " + PrimaryJobPosition.ToString + ""
                            TBL_Family.DefaultView.RowFilter = Filter
                            If TBL_Family.DefaultView.Count > 0 Then
                                For Each DV In TBL_Family.DefaultView
                                    Row_Family = DV.Row
                                    If Me.cboPeriod.Text = "All Data" Then
                                        tampil = True
                                    Else
                                        If Row_Family.DOC_NO <> "" Then
                                            If Me.cboPeriod.Text = "Contract Date" Then
                                                If Me.dtpFrom.Value.Date <= Row_Family._date And Me.dtpTo.Value.Date >= Row_Family._date Then
                                                    tampil = True
                                                Else
                                                    tampil = False
                                                End If
                                            ElseIf Me.cboPeriod.Text = "Contract End Date" Then
                                                If Me.dtpFrom.Value.Date <= Row_Family.start_date And Me.dtpTo.Value.Date >= Row_Family.start_date Then
                                                    tampil = True
                                                Else
                                                    tampil = False
                                                End If
                                            ElseIf Me.cboPeriod.Text = "Contract Start Date" Then
                                                If Me.dtpFrom.Value.Date <= Row_Family.end_date And Me.dtpTo.Value.Date >= Row_Family.end_date Then
                                                    tampil = True
                                                Else
                                                    tampil = False
                                                End If
                                            End If
                                        Else
                                            tampil = False
                                        End If
                                    End If
                                    If tampil = True Then
                                        If mListIdentifier.Contains(Row_Family.identifier) = False Then
                                            mListIdentifier.Add(Row_Family.identifier)
                                            Row_Dummy_Family = TBL_Dummy_Family.NewAnggota_KeluargaRow
                                            For a As Integer = 0 To TBL_Dummy_Family.Columns.Count - 1
                                                Row_Dummy_Family.Item(a) = Row_Family.Item(a)
                                            Next
                                            TBL_Dummy_Family.Rows.Add(Row_Dummy_Family)
                                        End If
                                    End If
                                Next
                            End If
                        Next
                        Filter = ""
                        For Each PrimaryDepartement In Me.ListDept
                            If Filter = "" Then
                                Filter = "[primary departemen] <> " + PrimaryDepartement.ToString + ""
                            Else
                                Filter += " AND [primary departemen] <> " + PrimaryDepartement.ToString + ""
                            End If
                        Next
                        TBL_Dummy_Family.DefaultView.RowFilter = Filter
                        If TBL_Dummy_Family.DefaultView.Count > 0 Then
                            While i <= TBL_Dummy_Family.DefaultView.Count - 1
                                Row_Dummy_Family = TBL_Dummy_Family.DefaultView(i).Row
                                TBL_Dummy_Family.Rows.Remove(Row_Dummy_Family)
                            End While
                        End If
                        Filter = ""
                        i = 0
                        For Each PrimaryDepartement In Me.ListEmp
                            If Filter = "" Then
                                Filter = "[primary] <> " + PrimaryDepartement.ToString + ""
                            Else
                                Filter += " AND [primary] <> " + PrimaryDepartement.ToString + ""
                            End If
                        Next
                        TBL_Dummy_Family.DefaultView.RowFilter = Filter
                        If TBL_Dummy_Family.DefaultView.Count > 0 Then
                            While i <= TBL_Dummy_Family.DefaultView.Count - 1
                                Row_Dummy_Family = TBL_Dummy_Family.DefaultView(i).Row
                                TBL_Dummy_Family.Rows.Remove(Row_Dummy_Family)
                            End While
                        End If
                        'Sorting data
                        TBL_Family.DefaultView.RowFilter = ""
                        TBL_Family.Rows.Clear()
                        TBL_Dummy_Family.DefaultView.RowFilter = ""
                        TBL_Dummy_Family.DefaultView.Sort = "[Dept Desc] ASC, [Nama Karyawan] ASC"
                        For Each DV In TBL_Dummy_Family.DefaultView
                            Row_Dummy_Family = DV.Row
                            Row_Family = TBL_Family.NewAnggota_KeluargaRow
                            For a As Integer = 0 To TBL_Family.Columns.Count - 1
                                Row_Family.Item(a) = Row_Dummy_Family.Item(a)
                            Next
                            TBL_Family.Rows.Add(Row_Family)
                        Next
                        'Else
                        '    If vChkJobPosition = True Then
                        '        For Each PrimaryJobPosition In mListJobPosition
                        '            TBL_Family.DefaultView.RowFilter = ""
                        '            Filter = "[primary position] = " + PrimaryJobPosition.ToString + ""
                        '            TBL_Family.DefaultView.RowFilter = Filter
                        '            If TBL_Family.DefaultView.Count > 0 Then
                        '                For Each DV In TBL_Family.DefaultView
                        '                    Row_Family = DV.Row
                        '                    If mListIdentifier.Contains(Row_Family.identifier) = False Then
                        '                        mListIdentifier.Add(Row_Family.identifier)
                        '                        Row_Dummy_Family = TBL_Dummy_Family.NewAnggota_KeluargaRow
                        '                        For a As Integer = 0 To TBL_Dummy_Family.Columns.Count - 1
                        '                            Row_Dummy_Family.Item(a) = Row_Family.Item(a)
                        '                        Next
                        '                        TBL_Dummy_Family.Rows.Add(Row_Dummy_Family)
                        '                    End If
                        '                Next
                        '            End If
                        '        Next
                        '        TBL_Family.DefaultView.RowFilter = ""
                        '        TBL_Family.Rows.Clear()
                        '        TBL_Dummy_Family.DefaultView.RowFilter = ""
                        '        TBL_Dummy_Family.DefaultView.Sort = "[Dept Desc] ASC, [Nama Karyawan] ASC"
                        '        For Each DV In TBL_Dummy_Family.DefaultView
                        '            Row_Dummy_Family = DV.Row
                        '            Row_Family = TBL_Family.NewAnggota_KeluargaRow
                        '            For a As Integer = 0 To TBL_Family.Columns.Count - 1
                        '                Row_Family.Item(a) = Row_Dummy_Family.Item(a)
                        '            Next
                        '            TBL_Family.Rows.Add(Row_Family)
                        '        Next
                        '    ElseIf vChkDepartement = True Then
                        '        For Each PrimaryDepartement In mListDepartement
                        '            TBL_Family.DefaultView.RowFilter = ""
                        '            Filter = "[primary departemen] = " + PrimaryDepartement.ToString + ""
                        '            TBL_Family.DefaultView.RowFilter = Filter
                        '            If TBL_Family.DefaultView.Count > 0 Then
                        '                For Each DV In TBL_Family.DefaultView
                        '                    Row_Family = DV.Row
                        '                    If mListIdentifier.Contains(Row_Family.identifier) = False Then
                        '                        mListIdentifier.Add(Row_Family.identifier)
                        '                        Row_Dummy_Family = TBL_Dummy_Family.NewAnggota_KeluargaRow
                        '                        For a As Integer = 0 To TBL_Dummy_Family.Columns.Count - 1
                        '                            Row_Dummy_Family.Item(a) = Row_Family.Item(a)
                        '                        Next
                        '                        TBL_Dummy_Family.Rows.Add(Row_Dummy_Family)
                        '                    End If
                        '                Next
                        '            End If
                        '        Next
                        '        TBL_Family.DefaultView.RowFilter = ""
                        '        TBL_Family.Rows.Clear()
                        '        TBL_Dummy_Family.DefaultView.RowFilter = ""
                        '        TBL_Dummy_Family.DefaultView.Sort = "[Dept Desc] ASC, [Nama Karyawan] ASC"
                        '        For Each DV In TBL_Dummy_Family.DefaultView
                        '            Row_Dummy_Family = DV.Row
                        '            Row_Family = TBL_Family.NewAnggota_KeluargaRow
                        '            For a As Integer = 0 To TBL_Family.Columns.Count - 1
                        '                Row_Family.Item(a) = Row_Dummy_Family.Item(a)
                        '            Next
                        '            TBL_Family.Rows.Add(Row_Family)
                        '        Next
                        '    End If
                    End If
                End If
            End If

            'Report Show Data
            If RboFormat1.Checked = True Then
                Report1.SetDataSource(DS)
                vCRV.ReportSource = Report1
                vCRV.Show()
            End If
            If RboFormat2.Checked = True Then
                If TBL_Format2.Rows.Count > 0 Then
                    For Each Row_Format2 In TBL_Format2.Rows
                        If Row_Format2.marital_status = Core_Module.ClsMasterKaryawan.EnumMartialStatus._Single Then
                            Row_Format2.Tanggal_Menikah_Str = "-"
                        Else
                            Row_Format2.Tanggal_Menikah_Str = Format(Row_Format2.Tanggal_Menikah, "dd MMM yyyy")
                        End If
                        Row_Format2.Golongan_Darah_Str = Me.FungsiGolDarah(Row_Format2.Golongan_Darah)
                    Next
                End If
                Report2.SetDataSource(DS)
                vCRV.ReportSource = Report2
                vCRV.Show()
            End If
            If RboFamilyData.Checked = True Then
                mListIdentifier.Clear()
                mListGroup.Clear()
                mListGroupDetail.Clear()
                Dim Nomor_Family As Integer = 0
                If TBL_Family.Rows.Count > 0 Then
                    For ij As Integer = 0 To TBL_Family.Rows.Count - 1
                        Row_Family = TBL_Family.Rows(ij)
                        'For Each Row_Family In TBL_Family.Rows
                        If mListGroup.Contains(Row_Family.primary_departemen.ToString) = False Then
                            mListGroup.Add(Row_Family.primary_departemen.ToString)
                            'Nomor_Family = 1
                            mListGroupDetail.Clear()
                            mListGroupDetail.Add(Row_Family.Primary.ToString)
                        Else
                            If mListGroupDetail.Contains(Row_Family.Primary.ToString) = False Then
                                mListGroupDetail.Add(Row_Family.Primary.ToString)
                                Row_Family.Visible_Garis_Penuh = 1
                                'Nomor_Family += 1
                            End If
                        End If
                        If mListIdentifier.Contains(Row_Family.Primary) = False Then
                            mListIdentifier.Add(Row_Family.Primary)
                            Nomor_Family += 1
                            Row_Family.Nomor = Nomor_Family
                            Row_Family.Visible = 1
                            Row_Family.Blood_Type_Str = Me.FungsiGolDarah(Row_Family.Blood_Type)
                        Else
                            Row_Family.Blood_Type_Str = Me.FungsiGolDarah(Row_Family.Blood_Type)
                        End If
                        Row_Family.Jenis_Kelamin_Str = Me.FungsiJenisKelamin(Row_Family.Jenis_Kelamin)
                        Row_Family.Hubungan_Keluarga_Str = Me.FungsiHubunganKeluarga(Row_Family.Hubungan_Keluarga)
                    Next
                End If
                Report3.SetDataSource(DS)
                vCRV.ReportSource = Report3
                vCRV.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA_Format1.Dispose()
            DTA_Format2.Dispose()
            DTA_Family.Dispose()
            TBL_Dummy_Format1.Dispose()
            TBL_Dummy_Format2.Dispose()
            TBL_Dummy_Family.Dispose()
            mListJobPosition = Nothing
            mListDepartement = Nothing
            mListPrimary = Nothing
            mListIdentifier = Nothing
            mListGroup = Nothing
            mListGroupDetail = Nothing
        End Try
    End Sub
    Private Function FilterAktifEmployee(ByRef vTbl As DataTable) As Boolean
        Dim vRow As DataRow
        Try
            If Me.CboFilterEmployee.SelectedIndex = 1 Then 'aktive
                vTbl.DefaultView.RowFilter = "[Aktif]=1"
                For i As Integer = vTbl.DefaultView.Count - 1 To 0 Step -1
                    vRow = vTbl.DefaultView(0).Row
                    vTbl.Rows.Remove(vRow)
                Next
            ElseIf Me.CboFilterEmployee.SelectedIndex = 2 Then 'not active
                vTbl.DefaultView.RowFilter = "[Aktif]=0"
                For i As Integer = vTbl.DefaultView.Count - 1 To 0 Step -1
                    vRow = vTbl.DefaultView(0).Row
                    vTbl.Rows.Remove(vRow)
                Next
            End If

            If Me.RboFormat1.Checked = True Then
                For Each vRow In vTbl.Rows
                    vRow.Item("str npwp category") = Me.GetStrCategoryNPWP(vRow.Item("npwp category"))
                Next
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function FungsiJenisKelamin(ByVal vEnum As Core_Module.ClsMasterKaryawan_Item.EnumSEx) As String
        Dim Retval As String = ""
        Try
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumSEx.Female Then
                Retval = "Female"
            Else
                Retval = "Male"
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function FungsiGolDarah(ByVal vEnum As Core_Module.ClsMasterKaryawan.EnumBloodType) As String
        Dim Retval As String = ""
        Try
            If vEnum = Core_Module.ClsMasterKaryawan.EnumBloodType.A Then
                Retval = "A"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan.EnumBloodType.B Then
                Retval = "B"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan.EnumBloodType.AB Then
                Retval = "AB"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan.EnumBloodType.O Then
                Retval = "O"
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function FungsiHubunganKeluarga(ByVal vEnum As Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga) As String
        Dim Retval As String = ""
        Try
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga.Husband Then
                Retval = "Husband"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga.Wife Then
                Retval = "Wife"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga.Son Then
                Retval = "Son"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga.Daughter Then
                Retval = "Daughter"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga.Brother Then
                Retval = "Brother"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga.Sister Then
                Retval = "Sister"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga.Father Then
                Retval = "Father"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga.Mother Then
                Retval = "Mother"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga.Uncle Then
                Retval = "Uncle"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga.Aunt Then
                Retval = "Aunt"
            End If
            If vEnum = Core_Module.ClsMasterKaryawan_Item.EnumHubKeluarga.Other Then
                Retval = "Other"
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function BatasanLaporan() As Boolean
      
        Try
            If Me.ChkJobPosition.Checked = True Then
                Me.ListJobPosition(ListJob, True)
                If Me.clbJobPosition.Items.Count > 0 Then
                    If ListJob.Count <= 0 Then
                        MsgBox("Please checked job position", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If
            If Me.ChkDepartement.Checked = True Then
                Me.ListDepartement(ListDept, True)
                If Me.ClbDepartement.Items.Count > 0 Then
                    If ListDept.Count <= 0 Then
                        MsgBox("Please checked department", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If
            If Me.ChkJobPosition.Checked = False And Me.ChkDepartement.Checked = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListJob = Nothing
            ListDept = Nothing
        End Try
    End Function

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        'Dim Parameter As Boolean
        Try
            'Parameter = Me.BatasanLaporan
            'Me.BuatLaporan(Me.crv1, Me.ChkJobPosition.Checked, Me.ChkDepartement.Checked, Parameter)
            Me.BuatLaporan(Me.crv1, Me.ChkJobPosition.Checked, Me.ChkDepartement.Checked, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Lbl_JB_Unchecked_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Lbl_JB_Unchecked.LinkClicked
        Try
            For i As Integer = 0 To Me.clbJobPosition.Items.Count - 1
                Me.clbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Lbl_JB_Checked_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Lbl_JB_Checked.LinkClicked
        Try
            For i As Integer = 0 To Me.clbJobPosition.Items.Count - 1
                Me.clbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Lbl_Dept_Unchecked_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Lbl_Dept_Unchecked.LinkClicked
        Try
            For i As Integer = 0 To Me.ClbDepartement.Items.Count - 1
                Me.ClbDepartement.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Lbl_Dept_Checked_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles Lbl_Dept_Checked.LinkClicked
        Try
            For i As Integer = 0 To Me.ClbDepartement.Items.Count - 1
                Me.ClbDepartement.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanEmployeeData_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.MinimumSize = New Drawing.Point(1000, 670)
            Me.KendalikanKontrol(EnumTipe.JobPosition = EnumTipe.JobPosition, True)
            Me.KendalikanKontrol(EnumTipe.Departement, True)
            Me.dgvEmployee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvDepartment.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvJobPosition.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.cboPeriod.SelectedIndex = 3
            Me.ChkJobPosition.Checked = True
            'Me.IsiListBoxJobPosition(Me.clbJobPosition, Me.TBLJobPosition)
            'For i As Integer = 0 To Me.clbJobPosition.Items.Count - 1
            '    Me.clbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            'Next
            Me.ChkDepartement.Checked = True
            'Me.IsiListBoxDepartement(Me.ClbDepartement, Me.TBLDepartement)
            'For i As Integer = 0 To Me.ClbDepartement.Items.Count - 1
            '    Me.ClbDepartement.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            'Next
            Me.CboFilterEmployee.SelectedIndex = 0
            Me.IsiSemuaTable()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub IsiSemuaTable()
        Dim DTADepartement As New DseLaporanEmployeeDataTableAdapters.tblmaster_departemenTableAdapter
        Dim DTAEmployee As New DseLaporanEmployeeDataTableAdapters.tblmaster_karyawanTableAdapter
        Dim DTAJob As New DseLaporanEmployeeDataTableAdapters.tblmaster_jobpositionTableAdapter
        Try
            DTADepartement.Fill(Me.TBLDepartement)
            DTAEmployee.Fill(Me.TBLEmployee)
            DTAJob.Fill(Me.TBLJobPosition)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub ChkJobPosition_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkJobPosition.CheckedChanged
        Try
            If Me.ChkJobPosition.Checked = True Then
                Me.KendalikanKontrol(EnumTipe.JobPosition, True)
            Else
                Me.KendalikanKontrol(EnumTipe.JobPosition, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkDepartement_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkDepartement.CheckedChanged
        Try
            If Me.ChkDepartement.Checked = True Then
                Me.KendalikanKontrol(EnumTipe.Departement, True)
            Else
                Me.KendalikanKontrol(EnumTipe.Departement, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvEmployee.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColEmployee_Code.Index Then
                    Me.Buffer = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColEmployee_Code.Index Then
                    Me.Buffer = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEmployee.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvEmployee.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryEmployee As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                Me.dgvEmployee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvEmployee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                If IsNothing(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColEmployee_Code.Index Then
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjEmp.ApakahSudahAda(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryEmployee = Me.ObjEmp.GetPrimaryByCode(Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemEmployee(PrimaryEmployee, e.RowIndex)
                            If index < 0 Then
                                Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, e.RowIndex).Value = PrimaryEmployee
                                Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(PrimaryEmployee)
                                Me.dgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(PrimaryEmployee)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexEmp = e.RowIndex
                                        BgwEmployee.RunWorkerAsync()
                                    Else
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterKaryawanCari
                            ij.txtCari.Text = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryEmploye <> -1 Then
                                index = Me.GetIndexItemEmployee(ij.PrimaryEmploye, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.dgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(ij.PrimaryEmploye)
                                Else
                                    If Me.Buffer = Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        Exit Sub
                                    End If
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")

                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexEmp = e.RowIndex
                                            BgwEmployee.RunWorkerAsync()
                                        Else
                                            Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                        Exit Sub
                                    Else
                                        Me.dgvEmployee.Item(Me.ColEmployee_primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                        Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value = Me.ObjEmp.GetCodeByPrimary(ij.PrimaryEmploye)
                                        Me.dgvEmployee.Item(Me.ColEmployee_Name.Index, e.RowIndex).Value = Me.ObjEmp.GetFullNameByPrimary(ij.PrimaryEmploye)
                                    End If
                                End If
                            ElseIf ij.PrimaryEmploye = -1 Then
                                If Me.dgvEmployee.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexEmp = e.RowIndex
                                        Me.BgwEmployee.RunWorkerAsync()
                                    Else
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                                If Me.dgvEmployee.Item(Me.ColEmployee_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexEmp = e.RowIndex
                                        Me.BgwEmployee.RunWorkerAsync()
                                    Else
                                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvEmployee)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvEmployee)
                    End If
                Else
                    If e.ColumnIndex = Me.ColEmployee_Code.Index Then
                        If IsNothing(Me.Buffer) = False Then
                            Me.IndexEmp = e.RowIndex
                            Me.BgwEmployee.RunWorkerAsync()
                        End If
                        Me.dgvEmployee.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvEmployee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvEmployee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Sub dgvDepartment_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvDepartment.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColDeptCode.Index Then
                    Me.Buffer = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColDeptCode.Index Then
                    Me.Buffer = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDepartment_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDepartment.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvDepartment.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryDept As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            Me.dgvDepartment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDepartment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            With row
                If IsNothing(Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColDeptCode.Index Then
                        Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjDept.ApakahSudahAda(Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryDept = Me.ObjDept.GetPrimary(Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemDepartment(PrimaryDept, e.RowIndex)
                            If index < 0 Then
                                Me.dgvDepartment.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value = PrimaryDept
                                Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(PrimaryDept)
                                Me.dgvDepartment.Item(Me.ColDeptDesc.Index, e.RowIndex).Value = Me.ObjDept.GetDeptDesc(PrimaryDept)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexDept = e.RowIndex
                                        Me.BgwDepartment.RunWorkerAsync()
                                    Else
                                        Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvDepartment)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterDepartementCari
                            ij.txtCari.Text = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryDepartement <> -1 Then
                                index = Me.GetIndexItemDepartment(ij.PrimaryDepartement, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvDepartment.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value = ij.PrimaryDepartement
                                    Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(ij.PrimaryDepartement)
                                    Me.dgvDepartment.Item(Me.ColDeptDesc.Index, e.RowIndex).Value = Me.ObjDept.GetDeptDesc(ij.PrimaryDepartement)
                                Else
                                    If Me.Buffer = Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        Exit Sub
                                    End If
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexDept = e.RowIndex
                                            Me.BgwDepartment.RunWorkerAsync()
                                        Else
                                            Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvDepartment)
                                        Exit Sub
                                    Else
                                        Me.dgvDepartment.Item(Me.ColDeptPrimary.Index, e.RowIndex).Value = ij.PrimaryDepartement
                                        Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value = Me.ObjDept.GetDeptCode(ij.PrimaryDepartement)
                                        Me.dgvDepartment.Item(Me.ColDeptDesc.Index, e.RowIndex).Value = Me.ObjDept.GetDeptDesc(ij.PrimaryDepartement)
                                    End If
                                End If
                            ElseIf ij.PrimaryDepartement = -1 Then
                                If Me.dgvDepartment.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexDept = e.RowIndex
                                        Me.BgwDepartment.RunWorkerAsync()
                                    Else
                                        Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvDepartment)
                                    Exit Sub
                                End If
                                If Me.dgvDepartment.Item(Me.ColDeptCode.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexDept = e.RowIndex
                                        Me.BgwDepartment.RunWorkerAsync()
                                    Else
                                        Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvDepartment)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvDepartment)
                    End If
                Else
                    If e.ColumnIndex = Me.ColDeptCode.Index Then
                        If IsNothing(Me.Buffer) = False Then
                            Me.IndexDept = e.RowIndex
                            Me.BgwDepartment.RunWorkerAsync()
                        End If
                        Me.dgvDepartment.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvDepartment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvDepartment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub
    Private Function GetIndexItemEmployee(ByVal vPrimary As Double, _
         ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvEmployee.Rows.Count - 1
                RowDGV = Me.dgvEmployee.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColEmployee_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Function GetIndexItemDepartment(ByVal vPrimary As Double, _
     ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvDepartment.Rows.Count - 1
                RowDGV = Me.dgvDepartment.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColDeptPrimary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Function GetIndexItemJobPosition(ByVal vPrimary As Double, _
         ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvJobPosition.Rows.Count - 1
                RowDGV = Me.dgvJobPosition.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColJobPosisitonPrimary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvJobPosition_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvJobPosition.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvJobPosition.Item(Me.ColDeptCode.Index, e.RowIndex).Value) = False Then
                If e.ColumnIndex = Me.ColDeptCode.Index Then
                    Me.Buffer = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColDeptCode.Index Then
                    Me.Buffer = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvJobPosition_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvJobPosition.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvJobPosition.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryJobPost As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            Me.dgvJobPosition.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvJobPosition.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            With row
                If IsNothing(Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.Buffer = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColJobPosisitonCode.Index Then
                        Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjJobPost.ApakahSudahAda(Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryJobPost = Me.ObjJobPost.GetPrimary(Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItemJobPosition(PrimaryJobPost, e.RowIndex)
                            If index < 0 Then
                                Me.dgvJobPosition.Item(Me.ColJobPosisitonPrimary.Index, e.RowIndex).Value = PrimaryJobPost
                                Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value = Me.ObjJobPost.GetCode(PrimaryJobPost)
                                Me.dgvJobPosition.Item(Me.ColJobPosisitonDesc.Index, e.RowIndex).Value = Me.ObjJobPost.GetDesc(PrimaryJobPost)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexJobPost = e.RowIndex
                                        Me.BgwJobPosition.RunWorkerAsync()
                                    Else
                                        Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvJobPosition)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterJobPosition_Cari
                            ij.txtCari.Text = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.mPrimary <> -1 Then
                                index = Me.GetIndexItemJobPosition(ij.mPrimary, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvJobPosition.Item(Me.ColJobPosisitonPrimary.Index, e.RowIndex).Value = ij.mPrimary
                                    Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value = Me.ObjJobPost.GetCode(ij.mPrimary)
                                    Me.dgvJobPosition.Item(Me.ColJobPosisitonDesc.Index, e.RowIndex).Value = Me.ObjJobPost.GetDesc(ij.mPrimary)
                                Else
                                    If Me.Buffer = Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value Then
                                        Exit Sub
                                    End If
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.IndexJobPost = e.RowIndex
                                            Me.BgwJobPosition.RunWorkerAsync()
                                        Else
                                            Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvJobPosition)
                                        Exit Sub
                                    Else
                                        Me.dgvJobPosition.Item(Me.ColJobPosisitonPrimary.Index, e.RowIndex).Value = ij.mPrimary
                                        Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value = Me.ObjJobPost.GetCode(ij.mPrimary)
                                        Me.dgvJobPosition.Item(Me.ColJobPosisitonDesc.Index, e.RowIndex).Value = Me.ObjJobPost.GetDesc(ij.mPrimary)
                                    End If
                                End If
                            ElseIf ij.mPrimary = -1 Then
                                If Me.dgvJobPosition.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexJobPost = e.RowIndex
                                        Me.BgwJobPosition.RunWorkerAsync()
                                    Else
                                        Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvJobPosition)
                                    Exit Sub
                                End If
                                If Me.dgvJobPosition.Item(Me.ColJobPosisitonCode.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.IndexJobPost = e.RowIndex
                                        Me.BgwJobPosition.RunWorkerAsync()
                                    Else
                                        Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvJobPosition)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvJobPosition)
                    End If
                Else
                    If e.ColumnIndex = Me.ColJobPosisitonCode.Index Then
                        If IsNothing(Me.Buffer) = False Then
                            Me.IndexJobPost = e.RowIndex
                            Me.BgwJobPosition.RunWorkerAsync()
                        End If
                        Me.dgvJobPosition.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvJobPosition.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvJobPosition.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        End Try
    End Sub

    Private Sub BgwEmployee_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwEmployee.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwEmployee_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwEmployee.RunWorkerCompleted
        Try
            If dgvEmployee.RowCount > 0 Then
                Me.dgvEmployee.Rows.RemoveAt(Me.IndexEmp)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDepartment_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwDepartment.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwDepartment_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwDepartment.RunWorkerCompleted
        Try
            If dgvDepartment.RowCount > 0 Then
                Me.dgvDepartment.Rows.RemoveAt(Me.IndexDept)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwJobPosition_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwJobPosition.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub BgwJobPosition_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwJobPosition.RunWorkerCompleted
        Try
            If dgvJobPosition.RowCount > 0 Then
                Me.dgvJobPosition.Rows.RemoveAt(Me.IndexJobPost)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvEmployee_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvEmployee.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvEmployee.Rows(Me.dgvEmployee.Rows.Count - 1)
            dr.Cells(Me.ColEmployee_primary.Index).Value = -1
            dr.Cells(Me.ColEmployee_Code.Index).Value = ""
            dr.Cells(Me.ColEmployee_Name.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDepartment_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvDepartment.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvDepartment.Rows(Me.dgvDepartment.Rows.Count - 1)
            dr.Cells(Me.ColDeptPrimary.Index).Value = -1
            dr.Cells(Me.ColDeptCode.Index).Value = ""
            dr.Cells(Me.ColDeptDesc.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvJobPosition_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvJobPosition.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvJobPosition.Rows(Me.dgvJobPosition.Rows.Count - 1)
            dr.Cells(Me.ColJobPosisitonPrimary.Index).Value = -1
            dr.Cells(Me.ColJobPosisitonCode.Index).Value = ""
            dr.Cells(Me.ColJobPosisitonDesc.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPeriod_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPeriod.SelectedValueChanged
        Try
            If Me.cboPeriod.Text = "All Data" Then
                Me.dtpFrom.Enabled = False
                Me.dtpTo.Enabled = False

            Else
                Me.dtpFrom.Enabled = True
                Me.dtpTo.Enabled = True

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetStrCategoryNPWP(ByVal vInt_Category As Integer) As String
        Dim vStr_Category As String = ""
        Try
            If vInt_Category = 0 Then
                vStr_Category = "TK0"
            ElseIf vInt_Category = 1 Then
                vStr_Category = "TK1"
            ElseIf vInt_Category = 2 Then
                vStr_Category = "TK2"
            ElseIf vInt_Category = 3 Then
                vStr_Category = "TK3"
            ElseIf vInt_Category = 4 Then
                vStr_Category = "K0"
            ElseIf vInt_Category = 5 Then
                vStr_Category = "K1"
            ElseIf vInt_Category = 6 Then
                vStr_Category = "K2"
            ElseIf vInt_Category = 7 Then
                vStr_Category = "K3"
            End If
            Return vStr_Category
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class