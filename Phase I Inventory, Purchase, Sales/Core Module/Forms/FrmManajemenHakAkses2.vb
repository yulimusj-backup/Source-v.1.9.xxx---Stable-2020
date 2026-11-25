Option Explicit On
Option Strict Off
Public Class FrmManajemenHakAkses2
    Public MainMenuStrip1 As MenuStrip
    ''Public TreeNode As TreeNode
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjHak As New ClsMasterUserGroup
    Dim ObjHak_Item As New ClsMasterUserGroupItem
    Dim baru As Boolean
    Dim TBL As New DSUser.tblmasterhakaksesDataTable
    Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter
    Dim Nd As Windows.Forms.TreeNode
    Public ObjCD As New ClsCekData
    Dim Aktif As Boolean = False
    Dim Berubah As Boolean = False
    Dim Collapsing As Boolean = True
    Dim CollapseNode As TreeNode
    Dim AllowCheckedChanged As Boolean = False
    Dim TBL1 As New DSUser.tblmasterhakaksesDataTable
    Dim row1 As DSUser.tblmasterhakaksesRow
    Dim ObjHakAkses As New Core_Module.ClsHakAkses

    Private Sub FrmManajemenHakAkses2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtGroupCode.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjHak.SimpanBaru(Me.trvAccess, Me.UserCode, Me.TBL1) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.ObjCD.CekVisibilitasMenu(Me.MainMenuStrip1, Me.UserCode)
                                Me.SalinKeField()
                                ''Me.NonAktifkanDetail()
                                ''Me.btnFind.Focus()
                                Me.btnOK.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjHak.SimpanPerubahan(Me.trvAccess, Me.UserCode, Me.TxtPrimaryGroup.Text, Me.TBL1) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.ObjCD.CekVisibilitasMenu(Me.MainMenuStrip1, Me.UserCode)
                                Me.SalinKeField()
                                ''Me.NonAktifkanDetail()
                                ''Me.btnEdit.Focus()
                                Me.btnOK.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmManajemenHakAkses2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'isikan tree view sesuai dengan menu strip yang diberikan
            DTA.Fill(TBL)
            Me.FillTree(MainMenuStrip1)
            Me.NonAktifkandetail()

            'kosongkan TextDescription
            Me.TxtAccesDesc.Clear()
            Me.dgvDetail.Rows.Clear()

            Me.ColDesc.SortMode = DataGridViewColumnSortMode.NotSortable
            Me.ColEnable.SortMode = DataGridViewColumnSortMode.NotSortable
            Me.ColNo.SortMode = DataGridViewColumnSortMode.NotSortable

            Me.dgvDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.SplitContainer1.SplitterDistance = 300
            Me.dgvDetail.RowHeadersVisible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FillDescription(ByVal TN As TreeNode, ByVal tulis As TextBox)
        Dim Obj As Object
        ''Dim tx As Windows.Forms.TreeNode
        ''Dim Tulis As Windows.Forms.TextBox
        Try
            ''Buat
            If Me.trvAccess.Nodes.Count <> -1 Then
                For Each Obj In TN.Nodes
                    TN = TryCast(Obj, Windows.Forms.TreeNode)
                    If IsNothing(TN) = False Then
                        If TN.Checked = True Then
                            tulis.Text = Me.GetDescription(TN)
                        End If
                    End If
                Next
            End If

            'buat node
            ''If Me.trvAccess.Nodes.Count <> -1 Then
            ''For Each TN In Me.trvAccess.Nodes
            ''If TN.Checked = True Then
            ''Me.GetDescription(TN)
            ''End If
            ''Next
            ''End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetDescription(ByVal i As TreeNode) As String
        '' Dim ds As New DseMasterUserGroup
        ''Dim row As DseMasterUserGroup.tblmasterhakaksesRow
        Dim DTA As New DseMasterUserGroupTableAdapters.tblmasterhakaksesTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '' For Each row In ds.tblmasterhakakses.Rows
            ''Me.TxtAccesDesc.Text = row.Description
            ''Next
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDescription(i.Name)) = False Then
                Return DTA.ScalarQuery_GetDescription(i.Name)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub FillTree(ByVal MS As MenuStrip)
        Dim Obj As Object
        Dim Mn As Windows.Forms.ToolStripMenuItem

        Try
            'kosongkan nodes
            Me.trvAccess.Nodes.Clear()

            'buat node
            For Each Obj In MS.Items
                Mn = TryCast(Obj, Windows.Forms.ToolStripMenuItem)
                If IsNothing(Mn) = False Then
                    Nd = New Windows.Forms.TreeNode
                    Me.trvAccess.Nodes.Add(Nd)
                    Nd.Text = Mn.Text
                    Nd.Name = Me.GetKodeHakAkses(Mn.Name.Substring(0, Mn.Name.Length - 1)) '--> kode disimpan disini

                    'kalo menu item ada anaknya, maka tambahkan juga anaknya ke dalam tree view
                    If Mn.DropDownItems.Count > 0 Then
                        Me.FillItem(Nd, Mn)
                    End If
                End If
            Next

            'Bersihkan tanda dan 
            For Each Nd In Me.trvAccess.Nodes
                Me.Bersihkan_Tanda_Dan(Nd)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FillItem(ByVal Nd As Windows.Forms.TreeNode, ByVal Mn As Windows.Forms.ToolStripMenuItem)
        Dim Obj As Object
        Dim MnItem As Windows.Forms.ToolStripMenuItem
        Dim NdItem As Windows.Forms.TreeNode
        Try
            For Each Obj In Mn.DropDownItems
                MnItem = TryCast(Obj, Windows.Forms.ToolStripMenuItem)
                If IsNothing(MnItem) = False Then
                    If MnItem.Name.Length <> 0 Then
                        NdItem = New Windows.Forms.TreeNode
                        NdItem.Text = MnItem.Text
                        NdItem.Name = Me.GetKodeHakAkses(MnItem.Name.Substring(0, MnItem.Name.Length - 1)) '--> kode disimpan disini
                        Nd.Nodes.Add(NdItem)
                        'kalo menu ada anaknya maka tambahkan juga anaknya ke dalam tree view
                        If MnItem.DropDownItems.Count > 0 Then
                            Me.FillItem(NdItem, MnItem)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Bersihkan_Tanda_Dan(ByVal ND As TreeNode)
        Dim ND1 As Windows.Forms.TreeNode
        Dim StartIndex, Index_Dan, Index_Dan2 As Integer
        Try
            StartIndex = 0
            Index_Dan = ND.Text.IndexOf("&", StartIndex)
            While Index_Dan <> -1 And StartIndex <= ND.Text.Length - 1
                Index_Dan2 = ND.Text.IndexOf("&&", StartIndex)
                If Index_Dan2 <> Index_Dan Then
                    'bukan tanda dan sesungguhnya
                    StartIndex = Index_Dan
                Else
                    StartIndex = Index_Dan + 1
                End If
                If Index_Dan = 0 Then
                    ND.Text = ND.Text.Substring(1)
                Else
                    ND.Text = ND.Text.Substring(0, Index_Dan) + ND.Text.Substring(Index_Dan + 1)
                End If
                Index_Dan = ND.Text.IndexOf("&", StartIndex)
            End While

            For Each ND1 In ND.Nodes
                Me.Bersihkan_Tanda_Dan(ND1)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''Tambahan untuk simpan node
    Private Function GetKodeHakAkses(ByVal vNamaHakAkses As String) As Double
        Dim Filter As String = ""
        Dim row As DSUser.tblmasterhakaksesRow

        Try
            Filter = "Nama ='" + vNamaHakAkses + "'"
            Me.TBL.DefaultView.RowFilter = Filter
            If Me.TBL.DefaultView.Count > 0 Then
                row = Me.TBL.DefaultView(0).Row
                Return row.Kode
            End If

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Aktifkandetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryGroup.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            'Me.trvAccess.Enabled = True  
            Me.trvAccess.BackColor = Color.White
            Me.TxtAccesDesc.BackColor = Me.TxtReadonly.BackColor
            Me.TxtAccesDesc.Enabled = True
            Me.Aktif = True
            Me.dgvDetail.ReadOnly = False
            Me.AllowCheckedChanged = True
            Me.chkSuperUser.Visible = True
            Me.chkSuperUser.TabStop = True
            If Me.ApakahSuperUser = True Then
                Me.chkSuperUser.Checked = True
            Else
                Me.chkSuperUser.Checked = False
            End If
            If Me.dgvDetail.RowCount > 0 Then
                Me.LinkLabel1.Enabled = True
                Me.LinkLabel2.Enabled = True
            Else
                Me.LinkLabel1.Enabled = False
                Me.LinkLabel2.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkandetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            'Me.trvAccess.Enabled = False  
            Me.TxtAccesDesc.Enabled = False
            Me.trvAccess.BackColor = Me.TxtReadonly.BackColor
            Me.TxtAccesDesc.BackColor = Me.TxtReadonly.BackColor
            Me.Aktif = False
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, True)
            Me.dgvDetail.ReadOnly = True
            Me.AllowCheckedChanged = False
            Me.chkSuperUser.Visible = False
            Me.chkSuperUser.TabStop = False
            Me.LinkLabel1.Enabled = False
            Me.LinkLabel2.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryGroup.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.TxtAccesDesc.Clear()
            Me.dgvDetail.Rows.Clear()

            For Each Me.row1 In Me.TBL1.Rows
                row1.Checked = 0
            Next
            For Each No As TreeNode In Me.trvAccess.Nodes
                No.Checked = False
                If No.Nodes.Count > 0 Then
                    Me.KosongkanCheck(No)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanCheck(ByVal No As TreeNode)
        Try
            If No.Nodes.Count = 0 Then
                No.Checked = False
            Else
                No.Checked = False
                For Each No1 As TreeNode In No.Nodes
                    Me.KosongkanCheck(No1)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariMasterHakAkses2

        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryLPB = -1 Then
                Me.AllowCheckedChanged = True
                If Me.ObjHak.CariDetail(ij.PrimaryLPB, Me.trvAccess, Me.TBL1) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
                Me.AllowCheckedChanged = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.TxtGroupCode.Focus()
            Me.TxtAccesDesc.Clear()
            Me.ObjHak.ResetObject(Me.trvAccess)
            Me.TxtGroupCode.Text = ""
            Me.Aktifkandetail()
            Me.KosongkanField()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KetentuanNode()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtGroupCode.Text = "" Then
                Me.NonAktifkandetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjHak.CariDetail(Me.TxtPrimaryGroup.Text, Me.trvAccess, Me.TBL1)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkandetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtGroupCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtGroupCode.Text <> "" Then
                If Me.ObjHak.ApakahBolehHapus(Me.TxtPrimaryGroup.Text, True, UserCode, True) = False Then
                    If Me.ObjHak.GetCreatedBy(Me.TxtPrimaryGroup.Text) = Me.UserCode Then
                        MsgBox("You cannot delete this data, there are user using this data.", _
                          MsgBoxStyle.Information, "Information")
                    End If
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjHak.Hapus(Me.TxtPrimaryGroup.Text, UserCode, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtGroupCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtGroupCode.Text <> "" Then
                If Me.ObjHak.ApakahBolehEdit(Me.ObjHak.Primary, UserCode, True, True) = True Then
                    Me.baru = False
                    Me.ObjHak.CariDetail(Me.ObjHak.Primary, Me.trvAccess, Me.TBL1)
                    Me.SalinKeField()
                    Me.TxtGroupCode.Focus()
                    Me.Aktifkandetail()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjHak.SimpanBaru(Me.trvAccess, Me.UserCode, Me.TBL1) Then
                            MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                            Me.ObjCD.CekVisibilitasMenu(Me.MainMenuStrip1, Me.UserCode)
                            Me.SalinKeField()
                            Me.NonAktifkandetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjHak.SimpanPerubahan(Me.trvAccess, Me.UserCode, CDbl(Me.TxtPrimaryGroup.Text), Me.TBL1) Then
                            MsgBox("Data saved.", MsgBoxStyle.Information, "Information")
                            Me.ObjCD.CekVisibilitasMenu(Me.MainMenuStrip1, Me.UserCode)
                            Me.SalinKeField()
                            Me.NonAktifkandetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("Code Group")
            KoleksiTextBox.Add(Me.TxtGroupCode)
            KoleksiNamaData.Add("Description")
            KoleksiTextBox.Add(Me.TxtDescription)

            If ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.ObjHak.TimeStamp <> Me.ObjHak.GetTimestamp(Me.ObjHak.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjHak.CariDetail(Me.ObjHak.Primary, Me.trvAccess, Me.TBL1)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjHak
                .GroupCode = Me.TxtGroupCode.Text
                .Description = Me.TxtDescription.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjHak

                Me.TxtPrimaryGroup.Text = .Primary
                Me.TxtGroupCode.Text = .GroupCode
                Me.TxtDescription.Text = .Description

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                Me.TxtAccesDesc.Clear()
                Me.dgvDetail.Rows.Clear()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoTransaksi_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGroupCode.GotFocus
        Try
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtGroupCode.Enabled = True Then
                    If Me.TxtGroupCode.ReadOnly = False Then
                        TxtGroupCode.BackColor = Me.BackColorAll
                    Else
                        TxtGroupCode.BackColor = Me.TxtReadonly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoTransaksi_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGroupCode.LostFocus
        Try
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtGroupCode.Enabled = True Then
                    If Me.TxtGroupCode.ReadOnly = False Then
                        TxtGroupCode.BackColor = Me.TxtPrimaryGroup.BackColor
                    Else
                        TxtGroupCode.BackColor = Me.TxtReadonly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCatatan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDescription.GotFocus
        Try
            ''If Me.TxtStockInNo.Enabled = True Then
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtDescription.ReadOnly = False Then
                    Me.TxtDescription.BackColor = Me.BackColorAll
                Else
                    TxtDescription.BackColor = Me.TxtReadonly.BackColor
                End If
            End If
            ''End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCatatan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDescription.LostFocus
        Try
            ''If Me.TxtStockInNo.Enabled = True Then
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtDescription.ReadOnly = False Then
                    TxtDescription.BackColor = Me.TxtPrimaryGroup.BackColor
                Else
                    TxtDescription.BackColor = Me.TxtReadonly.BackColor
                End If
            End If
            '' End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAcessDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccesDesc.GotFocus
        Try
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtAccesDesc.Enabled = True Then
                    If Me.TxtAccesDesc.ReadOnly = False Then
                        TxtAccesDesc.BackColor = Me.BackColorAll
                    Else
                        TxtAccesDesc.BackColor = Me.TxtReadonly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAcessDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccesDesc.LostFocus
        Try
            If Me.TxtDescription.Enabled = True Then
                If Me.TxtAccesDesc.Enabled = True Then
                    If Me.TxtAccesDesc.ReadOnly = False Then
                        TxtAccesDesc.BackColor = Me.TxtPrimaryGroup.BackColor
                    Else
                        TxtAccesDesc.BackColor = Me.TxtReadonly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub trvAccess_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trvAccess.AfterCheck
        Dim NamaAcuan As String = ""

        Try
            If Me.TxtGroupCode.Enabled = True Then
                NamaAcuan = Me.ObjHakAkses.GetNama(e.Node.Name)
                If NamaAcuan = "" Then
                Else
                    Me.TBL1.DefaultView.RowFilter = "[Nama Acuan] = '" + NamaAcuan + "'"
                    For Each DV As Data.DataRowView In Me.TBL1.DefaultView
                        Me.row1 = DV.Row
                        If e.Node.Checked = False Then
                            'semua hak akses turunannya di set false
                            row1.Checked = 0
                        Else
                            'semua hak akses turunannya diset sesuai nilai defaul
                            row1.Checked = row1.default_value_turunan
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub trvAccess_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trvAccess.AfterSelect
        Try
            If Me.TxtGroupCode.Enabled = True Then
                Me.TxtAccesDesc.Text = Me.GetDescription(e.Node)
                'If e.Node.Nodes.Count = 0 Then
                '    Me.IsikanHakAksesTurunan(e.Node.Name)
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub trvAccess_BeforeCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles trvAccess.BeforeCheck
        Try
            If Me.AllowCheckedChanged = False Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub trvAccess_BeforeCollapse(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles trvAccess.BeforeCollapse
        Me.Collapsing = False
        Me.CollapseNode = e.Node
    End Sub

    Private Sub trvAccess_BeforeExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewCancelEventArgs) Handles trvAccess.BeforeExpand
        Me.Collapsing = False
        Me.CollapseNode = e.Node
    End Sub

    Private Sub trvAccess_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles trvAccess.LostFocus
        Try
            If IsNothing(Me.trvAccess.SelectedNode) = True Then
                Me.TxtAccesDesc.Clear()
                Me.dgvDetail.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub trvAccess_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) _
    Handles trvAccess.NodeMouseClick
        Try
            Me.TxtAccesDesc.Text = Me.GetDescription(e.Node)
            If Me.Collapsing = False Then
                If e.Node.FullPath <> Me.CollapseNode.FullPath Then
                    Me.Collapsing = True
                End If
            End If

            If Me.AllowCheckedChanged = True Then
                If e.Node.Nodes.Count > 0 And Me.Collapsing = True Then
                    Me.Set_Check(e.Node, e.Node.Checked)
                End If
                Me.Collapsing = True

                Me.ObjHak_Item.Set_Status_Node(Me.trvAccess)
            End If

            If e.Node.Nodes.Count = 0 Then
                Me.IsikanHakAksesTurunan(e.Node.Name)
            Else
                Me.dgvDetail.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Set_Check(ByVal ND As TreeNode, ByVal Chk As Boolean)
        Dim ND_Child As TreeNode

        Try
            ND.Checked = Chk
            For Each ND_Child In ND.Nodes
                Me.Set_Check(ND_Child, Chk)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanHakAksesTurunan(ByVal vKodeHakAkses As Double)
        Dim NamaAcuan As String = ""
        Dim rowDGV As DataGridViewRow
        Dim ColWidth As Double = 0

        Try
            Me.dgvDetail.Rows.Clear()
            NamaAcuan = Me.ObjHakAkses.GetNama(vKodeHakAkses)
            If NamaAcuan = "" Then
            Else
                Me.TBL1.DefaultView.RowFilter = "[Nama Acuan] = '" + NamaAcuan + "'"
                For Each DV As Data.DataRowView In Me.TBL1.DefaultView
                    Me.dgvDetail.Rows.Add()
                    rowDGV = Me.dgvDetail.Rows(Me.dgvDetail.Rows.Count - 1)
                    Me.row1 = DV.Row
                    rowDGV.Cells(Me.ColDesc.Index).Value = Me.row1.Description
                    If row1.Checked = 1 Then
                        rowDGV.Cells(Me.ColEnable.Index).Value = Boolean.TrueString
                    Else
                        rowDGV.Cells(Me.ColEnable.Index).Value = Boolean.FalseString
                    End If
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvDetail.Rows.Count
                    rowDGV.Cells(Me.ColNamaAcuan.Index).Value = NamaAcuan
                Next

                'pastikan lebar kolom, klo ada scroll bar horizontal, fungsi ini jadinya ngaco 
                ColWidth = 0
                For Each Col1 As DataGridViewColumn In Me.dgvDetail.Columns
                    If Col1.Visible = True Then
                        ColWidth += Col1.Width
                    End If
                Next
                While Me.dgvDetail.Width < ColWidth + 4
                    Me.Width += 5
                End While

                If Me.dgvDetail.Rows.Count > 0 And Me.btnOK.Enabled = True Then
                    Me.LinkLabel1.Enabled = True
                    Me.LinkLabel2.Enabled = True
                Else
                    Me.LinkLabel1.Enabled = False
                    Me.LinkLabel2.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDetail_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvDetail.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColEnable.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDetail_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDetail.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColEnable.Index Then
                For Each Me.row1 In Me.TBL1.Rows
                    If Me.row1.Description = Me.dgvDetail.Item(Me.ColDesc.Index, e.RowIndex).Value.ToString And _
                    row1.Nama_Acuan = Me.dgvDetail.Item(Me.ColNamaAcuan.Index, e.RowIndex).Value.ToString Then
                        If Me.dgvDetail.Item(Me.ColEnable.Index, e.RowIndex).Value = Boolean.TrueString Then
                            Me.row1.Checked = 1
                        Else
                            Me.row1.Checked = 0
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahCheckSemua(ByVal Node As TreeNode) As Boolean
        Dim Node1 As TreeNode
        Try
            For Each Node1 In Node.Nodes
                If Node1.Nodes.Count > 0 Then
                    If Me.ApakahCheckSemua(Node1) = False Then
                        Return False
                    End If
                Else
                    'cek apakah ada kode hak akses yang di wakili oleh node ini
                    If Node1.Checked = False Then
                        If Me.ObjHakAkses.GetNama(Node1.Name) <> "" Then
                            Return False
                        End If
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahSuperUser() As Boolean
        'super user jika memiliki semua hak akses termasuk turunannya 
        Try
            'cek turunan
            For Each Me.row1 In Me.TBL1.Rows
                If row1.Checked = 0 Then
                    Return False
                End If
            Next

            'cek tree view
            For Each node As TreeNode In Me.trvAccess.Nodes
                If node.Nodes.Count > 0 Then
                    If Me.ApakahCheckSemua(node) = False Then
                        Return False
                    End If
                Else
                    'cek apakah ada kode hak akses yang di wakili oleh node ini
                    If node.Checked = False Then
                        If Me.ObjHakAkses.GetNama(node.Name) <> "" Then
                            Return False
                        End If
                    End If
                End If
            Next

            'kalo sampe sini maka return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SetSuperUser()
        Dim Node As TreeNode
        Try
            For Each Node In Me.trvAccess.Nodes
                Me.Set_Check(Node, True)
            Next
            For Each Me.row1 In Me.TBL1.Rows
                row1.Checked = 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSuperUser_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSuperUser.Leave
        Try
            If Me.TxtGroupCode.Enabled = True Then
                If Me.ApakahSuperUser = False And Me.chkSuperUser.Checked = True Then
                    If MsgBox("Set all access to this group?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.SetSuperUser()
                        If IsNothing(Me.trvAccess.SelectedNode) = False Then
                            Me.IsikanHakAksesTurunan(Me.trvAccess.SelectedNode.Name)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            If Me.dgvDetail.RowCount > 0 Then
                For Each rowDgv As DataGridViewRow In Me.dgvDetail.Rows
                    rowDgv.Cells(Me.ColEnable.Index).Value = Boolean.TrueString
                Next
                updateTBl()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            If Me.dgvDetail.RowCount > 0 Then
                For Each rowDgv As DataGridViewRow In Me.dgvDetail.Rows
                    rowDgv.Cells(Me.ColEnable.Index).Value = Boolean.FalseString
                Next
                updateTBl()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub updateTBl()
        Try
            If Me.dgvDetail.Rows.Count > 0 Then
                For Each rowDgv As DataGridViewRow In Me.dgvDetail.Rows
                    For Each Me.row1 In Me.TBL1.Rows
                        If Me.row1.Description = rowDgv.Cells(Me.ColDesc.Index).Value Then
                            If rowDgv.Cells(Me.ColEnable.Index).Value = Boolean.TrueString Then
                                Me.row1.Checked = 1
                            Else
                                Me.row1.Checked = 0
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub dgvDetail_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvDetail.DataError
        'do nothing
    End Sub
End Class
