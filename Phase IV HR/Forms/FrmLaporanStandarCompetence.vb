Option Explicit On
Option Strict Off

Imports System.Drawing
Imports System.Windows.Forms
Imports Core_Module
Public Class FrmLaporanStandarCompetence

    Dim TBLEvent As New DseLaporanStandarCompetence.tblmaster_jobpositionDataTable

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub IsiListBoxEvent(ByVal lbo As Windows.Forms.CheckedListBox, ByVal TBLEvent As DseLaporanStandarCompetence.tblmaster_jobpositionDataTable)
        Dim DTAEvent As New DseLaporanStandarCompetenceTableAdapters.tblmaster_jobpositionTableAdapter
        Try
            TBLEvent.Rows.Clear()
            DTAEvent.Fill(TBLEvent)
            lbo.DataSource = Nothing
            lbo.DataSource = TBLEvent.DefaultView

            'isikan No Event 
            lbo.DisplayMember = "Position"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAEvent.Dispose()
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblChecked.LinkClicked
        Try
            For i As Integer = 0 To Me.clbJobPosition.Items.Count - 1
                Me.clbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUnchecked.LinkClicked
        Try
            For i As Integer = 0 To Me.clbJobPosition.Items.Count - 1
                Me.clbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekClb() As Boolean
        Dim Ketemu As Boolean
        Try
            If Me.clbJobPosition.CheckedItems.Count > 0 Then
                Ketemu = True
            Else
                Ketemu = False
            End If
            Return Ketemu
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmLaporanStandarCompetence_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.MinimumSize = New Drawing.Point(1000, 670)
            Me.IsiListBoxEvent(Me.clbJobPosition, TBLEvent)
            For i As Integer = 0 To Me.clbJobPosition.Items.Count - 1
                Me.clbJobPosition.SetItemCheckState(i, Windows.Forms.CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Try
            If Me.CekClb = True Then
                Me.CetakLaporan(Me.crv1)
            Else
                MsgBox("Please select job position", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakLaporan(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer)
        Dim DV As Data.DataRowView
        Dim report As New CrpLaporanStandarCompetence
        Dim PrimaryJobPosition As Double = -1
        Dim ListJobPosition As New List(Of Double)
        Dim DTA As New DseLaporanStandarCompetenceTableAdapters.DataTable1TableAdapter
        Dim TBLDummy As New DseLaporanStandarCompetence.DataTable1DataTable
        Dim rowDummy As DseLaporanStandarCompetence.DataTable1Row
        Dim RowViewDummy As Data.DataRowView

        Dim DS As New DseLaporanStandarCompetence
        Dim TBL As DseLaporanStandarCompetence.DataTable1DataTable = DS.DataTable1
        Dim filter As String = ""
        Dim row As DseLaporanStandarCompetence.DataTable1Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim Number As Integer = 0
        Dim ListNumber As New List(Of Integer)
        Dim Visible As Integer = 0
        Dim PrimaryDummy As Double = -1
        Try
            TBLDummy.Rows.Clear()
            TBL.Rows.Clear()

            'Buat List data yang di cek list
            For Each DV In Me.clbJobPosition.CheckedItems
                PrimaryJobPosition = DV.Row.Item("Primary")
                If ListJobPosition.Contains(PrimaryJobPosition) = False Then
                    Number += 1
                    ListJobPosition.Add(PrimaryJobPosition)
                    ListNumber.Add(Number)
                End If
            Next

            'Datatable1 versi dummy di fill(..)
            DTA.Fill(TBLDummy)
            'Lakukan penghapusan row datatable1 untuk data yang tidak dichecklist, berdasarkan list job position
            For i As Integer = 0 To ListJobPosition.Count - 1
                PrimaryJobPosition = ListJobPosition.Item(i)
                filter = "[Primary Job Position] =" + PrimaryJobPosition.ToString + ""
                TBLDummy.DefaultView.RowFilter = filter
                TBLDummy.DefaultView.Sort = "[Competence] ASC"
                If TBLDummy.DefaultView.Count > 0 Then
                    Number = ListNumber.Item(i)
                    For Each RowViewDummy In TBLDummy.DefaultView
                        rowDummy = RowViewDummy.Row
                        row = TBL.NewDataTable1Row
                        row.Primary_Job_Position = rowDummy.Primary_Job_Position
                        row.Job_Name = rowDummy.Job_Name
                        row.Competence = rowDummy.Competence
                        row.value = rowDummy.value
                        row.Data_Type = rowDummy.Data_Type
                        row.Default_Precision = rowDummy.Default_Precision
                        row.Standar_Nilai = rowDummy.Standar_Nilai
                        row.Number = Number
                        If PrimaryDummy <> PrimaryJobPosition Then
                            row.Visible = 1
                        End If
                        PrimaryDummy = PrimaryJobPosition
                        TBL.AddDataTable1Row(row)
                    Next
                End If
            Next

            'lakukan konversi format berdasarkan data type
            For Each row In TBL.Rows
                If row.Data_Type = ClsMasterKompetensi.EnumDataType.Alfabetic Then
                    If row.value = ClsMasterKompetensi.Aphabetic.A Then
                        row.Standar_Nilai = "A"
                    End If
                    If row.value = ClsMasterKompetensi.Aphabetic.B Then
                        row.Standar_Nilai = "B"
                    End If
                    If row.value = ClsMasterKompetensi.Aphabetic.C Then
                        row.Standar_Nilai = "C"
                    End If
                    If row.value = ClsMasterKompetensi.Aphabetic.D Then
                        row.Standar_Nilai = "D"
                    End If
                    If row.value = ClsMasterKompetensi.Aphabetic.E Then
                        row.Standar_Nilai = "E"
                    End If
                    If row.value = ClsMasterKompetensi.Aphabetic.F Then
                        row.Standar_Nilai = "F"
                    End If
                End If
                If row.Data_Type = ClsMasterKompetensi.EnumDataType.Number Then
                    row.Standar_Nilai = ObjCD.DigitGrouping(row.value, row.Default_Precision)
                End If
                If row.Data_Type = ClsMasterKompetensi.EnumDataType.Tanggal Then
                    row.Standar_Nilai = Format(CDate(row.value), "dd MMM yyyy")
                End If
                If row.Data_Type = ClsMasterKompetensi.EnumDataType.Text Then
                    row.Standar_Nilai = row.value
                End If
                If row.Data_Type = ClsMasterKompetensi.EnumDataType.YN Then
                    If row.value = 0 Then
                        row.Standar_Nilai = "No"
                    End If
                    If row.value = 1 Then
                        row.Standar_Nilai = "Yes"
                    End If
                End If
            Next

            'Report
            report.SetDataSource(DS)
            vCRV.ReportSource = report
            vCRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListJobPosition = Nothing
            ListNumber = Nothing
            DS.Dispose()
            DTA.Dispose()
            TBLDummy.Dispose()
            ObjCD = Nothing
        End Try
    End Sub
End Class