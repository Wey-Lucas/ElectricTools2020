<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRegalConfigAddEdit
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
        Me.lblBlockName = New System.Windows.Forms.Label()
        Me.txtScaleXP = New System.Windows.Forms.TextBox()
        Me.lblScaleXP = New System.Windows.Forms.Label()
        Me.txtLines = New System.Windows.Forms.TextBox()
        Me.lblLines = New System.Windows.Forms.Label()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.chkAtivo = New System.Windows.Forms.CheckBox()
        Me.cmbBlockName = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'lblBlockName
        '
        Me.lblBlockName.AutoSize = True
        Me.lblBlockName.Location = New System.Drawing.Point(12, 9)
        Me.lblBlockName.Name = "lblBlockName"
        Me.lblBlockName.Size = New System.Drawing.Size(117, 13)
        Me.lblBlockName.TabIndex = 0
        Me.lblBlockName.Text = "Nome do bloco MODE:"
        '
        'txtScaleXP
        '
        Me.txtScaleXP.Location = New System.Drawing.Point(15, 73)
        Me.txtScaleXP.Name = "txtScaleXP"
        Me.txtScaleXP.Size = New System.Drawing.Size(138, 20)
        Me.txtScaleXP.TabIndex = 3
        '
        'lblScaleXP
        '
        Me.lblScaleXP.AutoSize = True
        Me.lblScaleXP.Location = New System.Drawing.Point(12, 54)
        Me.lblScaleXP.Name = "lblScaleXP"
        Me.lblScaleXP.Size = New System.Drawing.Size(123, 13)
        Me.lblScaleXP.TabIndex = 2
        Me.lblScaleXP.Text = "Escala da viewport (XP):"
        '
        'txtLines
        '
        Me.txtLines.Location = New System.Drawing.Point(15, 118)
        Me.txtLines.Name = "txtLines"
        Me.txtLines.Size = New System.Drawing.Size(138, 20)
        Me.txtLines.TabIndex = 5
        '
        'lblLines
        '
        Me.lblLines.AutoSize = True
        Me.lblLines.Location = New System.Drawing.Point(12, 99)
        Me.lblLines.Name = "lblLines"
        Me.lblLines.Size = New System.Drawing.Size(92, 13)
        Me.lblLines.TabIndex = 4
        Me.lblLines.Text = "Número de linhas:"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.Location = New System.Drawing.Point(200, 148)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancelar"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.Location = New System.Drawing.Point(281, 148)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 23)
        Me.cmdOK.TabIndex = 8
        Me.cmdOK.Text = "Aguardando..."
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'chkAtivo
        '
        Me.chkAtivo.AutoSize = True
        Me.chkAtivo.Location = New System.Drawing.Point(15, 151)
        Me.chkAtivo.Name = "chkAtivo"
        Me.chkAtivo.Size = New System.Drawing.Size(53, 17)
        Me.chkAtivo.TabIndex = 6
        Me.chkAtivo.Text = "Ativo."
        Me.chkAtivo.UseVisualStyleBackColor = True
        '
        'cmbBlockName
        '
        Me.cmbBlockName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbBlockName.FormattingEnabled = True
        Me.cmbBlockName.Location = New System.Drawing.Point(15, 28)
        Me.cmbBlockName.Name = "cmbBlockName"
        Me.cmbBlockName.Size = New System.Drawing.Size(341, 21)
        Me.cmbBlockName.TabIndex = 1
        '
        'frmRegalConfigAddEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(212, Byte), Integer), CType(CType(212, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(384, 181)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmbBlockName)
        Me.Controls.Add(Me.chkAtivo)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.txtLines)
        Me.Controls.Add(Me.lblLines)
        Me.Controls.Add(Me.txtScaleXP)
        Me.Controls.Add(Me.lblScaleXP)
        Me.Controls.Add(Me.lblBlockName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.MaximumSize = New System.Drawing.Size(800, 220)
        Me.MinimumSize = New System.Drawing.Size(400, 220)
        Me.Name = "frmRegalConfigAddEdit"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Aguardando..."
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblBlockName As Windows.Forms.Label
    Friend WithEvents txtScaleXP As Windows.Forms.TextBox
    Friend WithEvents lblScaleXP As Windows.Forms.Label
    Friend WithEvents txtLines As Windows.Forms.TextBox
    Friend WithEvents lblLines As Windows.Forms.Label
    Friend WithEvents cmdCancel As Windows.Forms.Button
    Friend WithEvents cmdOK As Windows.Forms.Button
    Friend WithEvents chkAtivo As Windows.Forms.CheckBox
    Friend WithEvents cmbBlockName As Windows.Forms.ComboBox
End Class
