<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRegalConfig
    Inherits System.Windows.Forms.Form

    'Descartar substituições de formulário para limpar a lista de componentes.
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

    'Exigido pelo Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'OBSERVAÇÃO: o procedimento a seguir é exigido pelo Windows Form Designer
    'Pode ser modificado usando o Windows Form Designer.  
    'Não o modifique usando o editor de códigos.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.lblGeral = New System.Windows.Forms.Label()
        Me.dgvGeral = New System.Windows.Forms.DataGridView()
        Me.cmdAdicionar = New System.Windows.Forms.Button()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdRemove = New System.Windows.Forms.Button()
        Me.cmdEdit = New System.Windows.Forms.Button()
        Me.btnEditComplementaryBlock = New System.Windows.Forms.Button()
        Me.btnRemoveComplementaryBlock = New System.Windows.Forms.Button()
        Me.btnAddComplementaryBlock = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.lblAlimBlockName = New System.Windows.Forms.Label()
        Me.txtAlimBlockName = New System.Windows.Forms.TextBox()
        Me.lblComplementaryBlock = New System.Windows.Forms.Label()
        Me.dgvComplementaryBlocks = New System.Windows.Forms.DataGridView()
        CType(Me.dgvGeral, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvComplementaryBlocks, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblGeral
        '
        Me.lblGeral.AutoSize = True
        Me.lblGeral.Location = New System.Drawing.Point(12, 51)
        Me.lblGeral.Name = "lblGeral"
        Me.lblGeral.Size = New System.Drawing.Size(147, 13)
        Me.lblGeral.TabIndex = 2
        Me.lblGeral.Text = "Configurações blocos MODE:"
        '
        'dgvGeral
        '
        Me.dgvGeral.AllowUserToAddRows = False
        Me.dgvGeral.AllowUserToDeleteRows = False
        Me.dgvGeral.AllowUserToOrderColumns = True
        Me.dgvGeral.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvGeral.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvGeral.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvGeral.BackgroundColor = System.Drawing.SystemColors.ButtonFace
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGeral.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.dgvGeral.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.Color.Silver
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.Color.White
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvGeral.DefaultCellStyle = DataGridViewCellStyle6
        Me.dgvGeral.Location = New System.Drawing.Point(15, 70)
        Me.dgvGeral.MultiSelect = False
        Me.dgvGeral.Name = "dgvGeral"
        Me.dgvGeral.ReadOnly = True
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.Color.Silver
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.Color.White
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGeral.RowHeadersDefaultCellStyle = DataGridViewCellStyle7
        DataGridViewCellStyle8.BackColor = System.Drawing.Color.Silver
        DataGridViewCellStyle8.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.Color.White
        Me.dgvGeral.RowsDefaultCellStyle = DataGridViewCellStyle8
        Me.dgvGeral.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvGeral.Size = New System.Drawing.Size(427, 119)
        Me.dgvGeral.TabIndex = 3
        '
        'cmdAdicionar
        '
        Me.cmdAdicionar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAdicionar.FlatAppearance.BorderSize = 0
        Me.cmdAdicionar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdAdicionar.Image = Global.ElectricTools2020.My.Resources.Resources.ADICIONAR2
        Me.cmdAdicionar.Location = New System.Drawing.Point(448, 70)
        Me.cmdAdicionar.Name = "cmdAdicionar"
        Me.cmdAdicionar.Size = New System.Drawing.Size(25, 25)
        Me.cmdAdicionar.TabIndex = 4
        Me.ToolTip.SetToolTip(Me.cmdAdicionar, "Adicionar.")
        Me.cmdAdicionar.UseVisualStyleBackColor = True
        '
        'cmdRemove
        '
        Me.cmdRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdRemove.FlatAppearance.BorderSize = 0
        Me.cmdRemove.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdRemove.Image = Global.ElectricTools2020.My.Resources.Resources.REMOVER2
        Me.cmdRemove.Location = New System.Drawing.Point(448, 100)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(25, 25)
        Me.cmdRemove.TabIndex = 5
        Me.ToolTip.SetToolTip(Me.cmdRemove, "Remover.")
        Me.cmdRemove.UseVisualStyleBackColor = True
        '
        'cmdEdit
        '
        Me.cmdEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdit.FlatAppearance.BorderSize = 0
        Me.cmdEdit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdEdit.Image = Global.ElectricTools2020.My.Resources.Resources.EDITAR2
        Me.cmdEdit.Location = New System.Drawing.Point(448, 130)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(25, 25)
        Me.cmdEdit.TabIndex = 6
        Me.ToolTip.SetToolTip(Me.cmdEdit, "Editar.")
        Me.cmdEdit.UseVisualStyleBackColor = True
        '
        'btnEditComplementaryBlock
        '
        Me.btnEditComplementaryBlock.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnEditComplementaryBlock.FlatAppearance.BorderSize = 0
        Me.btnEditComplementaryBlock.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnEditComplementaryBlock.Image = Global.ElectricTools2020.My.Resources.Resources.EDITAR2
        Me.btnEditComplementaryBlock.Location = New System.Drawing.Point(448, 277)
        Me.btnEditComplementaryBlock.Name = "btnEditComplementaryBlock"
        Me.btnEditComplementaryBlock.Size = New System.Drawing.Size(25, 25)
        Me.btnEditComplementaryBlock.TabIndex = 12
        Me.ToolTip.SetToolTip(Me.btnEditComplementaryBlock, "Editar.")
        Me.btnEditComplementaryBlock.UseVisualStyleBackColor = True
        '
        'btnRemoveComplementaryBlock
        '
        Me.btnRemoveComplementaryBlock.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRemoveComplementaryBlock.FlatAppearance.BorderSize = 0
        Me.btnRemoveComplementaryBlock.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRemoveComplementaryBlock.Image = Global.ElectricTools2020.My.Resources.Resources.REMOVER2
        Me.btnRemoveComplementaryBlock.Location = New System.Drawing.Point(448, 247)
        Me.btnRemoveComplementaryBlock.Name = "btnRemoveComplementaryBlock"
        Me.btnRemoveComplementaryBlock.Size = New System.Drawing.Size(25, 25)
        Me.btnRemoveComplementaryBlock.TabIndex = 11
        Me.ToolTip.SetToolTip(Me.btnRemoveComplementaryBlock, "Remover.")
        Me.btnRemoveComplementaryBlock.UseVisualStyleBackColor = True
        '
        'btnAddComplementaryBlock
        '
        Me.btnAddComplementaryBlock.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAddComplementaryBlock.FlatAppearance.BorderSize = 0
        Me.btnAddComplementaryBlock.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAddComplementaryBlock.Image = Global.ElectricTools2020.My.Resources.Resources.ADICIONAR2
        Me.btnAddComplementaryBlock.Location = New System.Drawing.Point(448, 217)
        Me.btnAddComplementaryBlock.Name = "btnAddComplementaryBlock"
        Me.btnAddComplementaryBlock.Size = New System.Drawing.Size(25, 25)
        Me.btnAddComplementaryBlock.TabIndex = 10
        Me.ToolTip.SetToolTip(Me.btnAddComplementaryBlock, "Adicionar.")
        Me.btnAddComplementaryBlock.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(397, 349)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Fechar"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'lblAlimBlockName
        '
        Me.lblAlimBlockName.AutoSize = True
        Me.lblAlimBlockName.Location = New System.Drawing.Point(12, 9)
        Me.lblAlimBlockName.Name = "lblAlimBlockName"
        Me.lblAlimBlockName.Size = New System.Drawing.Size(240, 13)
        Me.lblAlimBlockName.TabIndex = 0
        Me.lblAlimBlockName.Text = "Nome ou parte do nome do do bloco alimentador:"
        '
        'txtAlimBlockName
        '
        Me.txtAlimBlockName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAlimBlockName.Location = New System.Drawing.Point(15, 25)
        Me.txtAlimBlockName.Name = "txtAlimBlockName"
        Me.txtAlimBlockName.Size = New System.Drawing.Size(457, 20)
        Me.txtAlimBlockName.TabIndex = 1
        '
        'lblComplementaryBlock
        '
        Me.lblComplementaryBlock.AutoSize = True
        Me.lblComplementaryBlock.Location = New System.Drawing.Point(12, 195)
        Me.lblComplementaryBlock.Name = "lblComplementaryBlock"
        Me.lblComplementaryBlock.Size = New System.Drawing.Size(208, 13)
        Me.lblComplementaryBlock.TabIndex = 8
        Me.lblComplementaryBlock.Text = "Configurações de blocos Complementares:"
        '
        'dgvComplementaryBlocks
        '
        Me.dgvComplementaryBlocks.BackgroundColor = System.Drawing.SystemColors.ButtonFace
        Me.dgvComplementaryBlocks.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvComplementaryBlocks.Location = New System.Drawing.Point(15, 214)
        Me.dgvComplementaryBlocks.Name = "dgvComplementaryBlocks"
        Me.dgvComplementaryBlocks.Size = New System.Drawing.Size(427, 126)
        Me.dgvComplementaryBlocks.TabIndex = 9
        '
        'frmRegalConfig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(212, Byte), Integer), CType(CType(212, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(484, 384)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnEditComplementaryBlock)
        Me.Controls.Add(Me.btnRemoveComplementaryBlock)
        Me.Controls.Add(Me.btnAddComplementaryBlock)
        Me.Controls.Add(Me.dgvComplementaryBlocks)
        Me.Controls.Add(Me.lblComplementaryBlock)
        Me.Controls.Add(Me.txtAlimBlockName)
        Me.Controls.Add(Me.lblAlimBlockName)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdEdit)
        Me.Controls.Add(Me.cmdRemove)
        Me.Controls.Add(Me.cmdAdicionar)
        Me.Controls.Add(Me.dgvGeral)
        Me.Controls.Add(Me.lblGeral)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.MinimumSize = New System.Drawing.Size(500, 400)
        Me.Name = "frmRegalConfig"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Configurações do registro de alimentadores"
        CType(Me.dgvGeral, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvComplementaryBlocks, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblGeral As Windows.Forms.Label
    Friend WithEvents dgvGeral As Windows.Forms.DataGridView
    Friend WithEvents cmdAdicionar As Windows.Forms.Button
    Friend WithEvents ToolTip As Windows.Forms.ToolTip
    Friend WithEvents cmdRemove As Windows.Forms.Button
    Friend WithEvents cmdEdit As Windows.Forms.Button
    Friend WithEvents cmdCancel As Windows.Forms.Button
    Friend WithEvents lblAlimBlockName As Windows.Forms.Label
    Friend WithEvents txtAlimBlockName As Windows.Forms.TextBox
    Friend WithEvents lblComplementaryBlock As Windows.Forms.Label
    Friend WithEvents dgvComplementaryBlocks As Windows.Forms.DataGridView
    Friend WithEvents btnEditComplementaryBlock As Windows.Forms.Button
    Friend WithEvents btnRemoveComplementaryBlock As Windows.Forms.Button
    Friend WithEvents btnAddComplementaryBlock As Windows.Forms.Button
End Class
