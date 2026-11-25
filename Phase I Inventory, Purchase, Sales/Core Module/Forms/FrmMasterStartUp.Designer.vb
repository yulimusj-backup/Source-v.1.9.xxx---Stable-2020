<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterStartUp
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkDeleteUnit = New System.Windows.Forms.CheckBox
        Me.chkDeleteCustomer = New System.Windows.Forms.CheckBox
        Me.chkDeleteProductGroup = New System.Windows.Forms.CheckBox
        Me.chkDeleteBudget = New System.Windows.Forms.CheckBox
        Me.chkDeleteCOA = New System.Windows.Forms.CheckBox
        Me.chkDeleteUser = New System.Windows.Forms.CheckBox
        Me.chkDeleteSupplier = New System.Windows.Forms.CheckBox
        Me.ChkDeleteWarehouse = New System.Windows.Forms.CheckBox
        Me.chkDeleteProduct = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.ChkDeleteEmail_Activities = New System.Windows.Forms.CheckBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnDelete)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.Controls.Add(Me.CheckBox2)
        Me.GroupBox1.Controls.Add(Me.ChkDeleteEmail_Activities)
        Me.GroupBox1.Controls.Add(Me.chkDeleteSupplier)
        Me.GroupBox1.Controls.Add(Me.ChkDeleteWarehouse)
        Me.GroupBox1.Controls.Add(Me.chkDeleteProduct)
        Me.GroupBox1.Controls.Add(Me.chkDeleteUnit)
        Me.GroupBox1.Controls.Add(Me.chkDeleteCustomer)
        Me.GroupBox1.Controls.Add(Me.chkDeleteProductGroup)
        Me.GroupBox1.Controls.Add(Me.chkDeleteBudget)
        Me.GroupBox1.Controls.Add(Me.chkDeleteCOA)
        Me.GroupBox1.Controls.Add(Me.chkDeleteUser)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(653, 440)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Data Delete"
        '
        'chkDeleteUnit
        '
        Me.chkDeleteUnit.AutoSize = True
        Me.chkDeleteUnit.Location = New System.Drawing.Point(6, 134)
        Me.chkDeleteUnit.Name = "chkDeleteUnit"
        Me.chkDeleteUnit.Size = New System.Drawing.Size(45, 17)
        Me.chkDeleteUnit.TabIndex = 6
        Me.chkDeleteUnit.Text = "Unit"
        Me.chkDeleteUnit.UseVisualStyleBackColor = True
        '
        'chkDeleteCustomer
        '
        Me.chkDeleteCustomer.AutoSize = True
        Me.chkDeleteCustomer.Location = New System.Drawing.Point(6, 111)
        Me.chkDeleteCustomer.Name = "chkDeleteCustomer"
        Me.chkDeleteCustomer.Size = New System.Drawing.Size(85, 17)
        Me.chkDeleteCustomer.TabIndex = 5
        Me.chkDeleteCustomer.Text = "Customer ++"
        Me.chkDeleteCustomer.UseVisualStyleBackColor = True
        '
        'chkDeleteProductGroup
        '
        Me.chkDeleteProductGroup.AutoSize = True
        Me.chkDeleteProductGroup.Location = New System.Drawing.Point(6, 88)
        Me.chkDeleteProductGroup.Name = "chkDeleteProductGroup"
        Me.chkDeleteProductGroup.Size = New System.Drawing.Size(233, 17)
        Me.chkDeleteProductGroup.TabIndex = 4
        Me.chkDeleteProductGroup.Text = "Product Group, Brand, Sub Brand, Category"
        Me.chkDeleteProductGroup.UseVisualStyleBackColor = True
        '
        'chkDeleteBudget
        '
        Me.chkDeleteBudget.AutoSize = True
        Me.chkDeleteBudget.Location = New System.Drawing.Point(6, 65)
        Me.chkDeleteBudget.Name = "chkDeleteBudget"
        Me.chkDeleteBudget.Size = New System.Drawing.Size(60, 17)
        Me.chkDeleteBudget.TabIndex = 3
        Me.chkDeleteBudget.Text = "Budget"
        Me.chkDeleteBudget.UseVisualStyleBackColor = True
        '
        'chkDeleteCOA
        '
        Me.chkDeleteCOA.AutoSize = True
        Me.chkDeleteCOA.Location = New System.Drawing.Point(6, 42)
        Me.chkDeleteCOA.Name = "chkDeleteCOA"
        Me.chkDeleteCOA.Size = New System.Drawing.Size(108, 17)
        Me.chkDeleteCOA.TabIndex = 2
        Me.chkDeleteCOA.Text = "Chart Of Account"
        Me.chkDeleteCOA.UseVisualStyleBackColor = True
        '
        'chkDeleteUser
        '
        Me.chkDeleteUser.AutoSize = True
        Me.chkDeleteUser.Location = New System.Drawing.Point(6, 19)
        Me.chkDeleteUser.Name = "chkDeleteUser"
        Me.chkDeleteUser.Size = New System.Drawing.Size(95, 17)
        Me.chkDeleteUser.TabIndex = 1
        Me.chkDeleteUser.Text = "User + Access"
        Me.chkDeleteUser.UseVisualStyleBackColor = True
        '
        'chkDeleteSupplier
        '
        Me.chkDeleteSupplier.AutoSize = True
        Me.chkDeleteSupplier.Location = New System.Drawing.Point(6, 203)
        Me.chkDeleteSupplier.Name = "chkDeleteSupplier"
        Me.chkDeleteSupplier.Size = New System.Drawing.Size(64, 17)
        Me.chkDeleteSupplier.TabIndex = 9
        Me.chkDeleteSupplier.Text = "Supplier"
        Me.chkDeleteSupplier.UseVisualStyleBackColor = True
        '
        'ChkDeleteWarehouse
        '
        Me.ChkDeleteWarehouse.AutoSize = True
        Me.ChkDeleteWarehouse.Location = New System.Drawing.Point(6, 180)
        Me.ChkDeleteWarehouse.Name = "ChkDeleteWarehouse"
        Me.ChkDeleteWarehouse.Size = New System.Drawing.Size(155, 17)
        Me.ChkDeleteWarehouse.TabIndex = 8
        Me.ChkDeleteWarehouse.Text = "Warehouse && Storage Area"
        Me.ChkDeleteWarehouse.UseVisualStyleBackColor = True
        '
        'chkDeleteProduct
        '
        Me.chkDeleteProduct.AutoSize = True
        Me.chkDeleteProduct.Location = New System.Drawing.Point(6, 157)
        Me.chkDeleteProduct.Name = "chkDeleteProduct"
        Me.chkDeleteProduct.Size = New System.Drawing.Size(63, 17)
        Me.chkDeleteProduct.TabIndex = 7
        Me.chkDeleteProduct.Text = "Product"
        Me.chkDeleteProduct.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(6, 272)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(64, 17)
        Me.CheckBox1.TabIndex = 12
        Me.CheckBox1.Text = "Supplier"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(6, 249)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(155, 17)
        Me.CheckBox2.TabIndex = 11
        Me.CheckBox2.Text = "Warehouse && Storage Area"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'ChkDeleteEmail_Activities
        '
        Me.ChkDeleteEmail_Activities.AutoSize = True
        Me.ChkDeleteEmail_Activities.Location = New System.Drawing.Point(6, 226)
        Me.ChkDeleteEmail_Activities.Name = "ChkDeleteEmail_Activities"
        Me.ChkDeleteEmail_Activities.Size = New System.Drawing.Size(105, 17)
        Me.ChkDeleteEmail_Activities.TabIndex = 10
        Me.ChkDeleteEmail_Activities.Text = "Email && Activities"
        Me.ChkDeleteEmail_Activities.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(562, 13)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 13
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(667, 472)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(659, 446)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Data Delete"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(659, 446)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Data 2"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 6)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(643, 52)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Qty Begining Balance"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(6, 19)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(94, 23)
        Me.Button1.TabIndex = 14
        Me.Button1.Text = "Delete Begining Balance"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmMasterStartUp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(667, 472)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "FrmMasterStartUp"
        Me.Text = "FrmMasterStartUp"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkDeleteUnit As System.Windows.Forms.CheckBox
    Friend WithEvents chkDeleteCustomer As System.Windows.Forms.CheckBox
    Friend WithEvents chkDeleteProductGroup As System.Windows.Forms.CheckBox
    Friend WithEvents chkDeleteBudget As System.Windows.Forms.CheckBox
    Friend WithEvents chkDeleteCOA As System.Windows.Forms.CheckBox
    Friend WithEvents chkDeleteUser As System.Windows.Forms.CheckBox
    Friend WithEvents chkDeleteSupplier As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDeleteWarehouse As System.Windows.Forms.CheckBox
    Friend WithEvents chkDeleteProduct As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDeleteEmail_Activities As System.Windows.Forms.CheckBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
