Option Explicit On
Option Strict Off
Imports Core_Module, System.Drawing, System.Windows.Forms

Public Class WaitForm
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Result As Integer = 0
        Dim DTA As New DataSet1TableAdapters.DataTable3TableAdapter
        Dim TBL As New DataSet1.DataTable3DataTable
        Dim Row1 As DataSet1.DataTable3Row
        Dim Row2 As DataSet1.DataTable3Row
        Dim Idx1 As Double = 0
        Dim Idx2 As Double = 0
        Dim ListDummy As New List(Of Double)
        Dim ListCodeEmployee As New List(Of String)
        Dim Code As String
        Dim Filter As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                While Idx1 <= TBL.Rows.Count - 1
                    Row1 = TBL.Rows(Idx1)
                    If Idx1 = TBL.Rows.Count - 1 Then
                        Exit While
                    Else
                        Idx2 += 1
                        Row2 = TBL.Rows(Idx2)
                        If Idx1 = 0 Then
                            ListDummy.Add(Row2.primary_entity)
                        End If
                        If ListDummy.Contains(Row2.primary_entity) = False Then
                            ListDummy.Add(Row2.primary_entity)
                        Else
                            'Lakukan pengecekan data allocate workshift
                            If Row2.start_date <= Row1.end_date Then
                                Result += 1
                                If ListCodeEmployee.Contains(Row1.primary_entity) = False Then
                                    ListCodeEmployee.Add(Row1.primary_entity)
                                End If
                            End If
                        End If
                        Idx1 += 1
                    End If
                End While
            End If
            Me.DataGridView1.DataSource = Nothing
            Me.DataGridView1.Rows.Clear()
            If ListCodeEmployee.Count > 0 Then
                For Each Code In ListCodeEmployee
                    If Filter = "" Then
                        Filter = "[primary entity] = " + Code.ToString + ""
                    Else
                        Filter += " or [primary entity] = " + Code.ToString + ""
                    End If
                Next
                TBL.DefaultView.RowFilter = Filter
                TBL.DefaultView.Sort = "[Emp_Code], [start date]"
                Me.DataGridView1.DataSource = TBL.DefaultView
            End If
            MsgBox("Hasil baris error = " + Result.ToString, MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Information")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ListCodeEmployee = Nothing
            ListDummy = Nothing
        End Try
    End Sub

    Private Sub WaitForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Information")
        End Try
    End Sub
End Class