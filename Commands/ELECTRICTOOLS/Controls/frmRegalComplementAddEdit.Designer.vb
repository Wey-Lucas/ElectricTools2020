<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRegalComplementAddEdit
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
        Me.lblComplementBlockName = New System.Windows.Forms.Label()
        Me.lblScaleXP = New System.Windows.Forms.Label()
        Me.cmbBlockName = New System.Windows.Forms.ComboBox()
        Me.txtScaleXP = New System.Windows.Forms.TextBox()
        Me.chkAtivo = New System.Windows.Forms.CheckBox()
        Me.BtnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblComplementBlockName
        '
        Me.lblComplementBlockName.AutoSize = True
        Me.lblComplementBlockName.Location = New System.Drawing.Point(12, 9)
        Me.lblComplementBlockName.Name = "lblComplementBlockName"
        Me.lblComplementBlockName.Size = New System.Drawing.Size(153, 13)
        Me.lblComplementBlockName.TabIndex = 0
        Me.lblComplementBlockName.Text = "Nome do Bloco Complementar:"
        '
        'lblScaleXP
        '
        Me.lblScaleXP.AutoSize = True
        Me.lblScaleXP.Location = New System.Drawing.Point(12, 54)
        Me.lblScaleXP.Name = "lblScaleXP"
        Me.lblScaleXP.Size = New System.Drawing.Size(123, 13)
        Me.lblScaleXP.TabIndex = 1
        Me.lblScaleXP.Text = "Escala da viewport (XP):"
        '
        'cmbBlockName
        '
        Me.cmbBlockName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbBlockName.FormattingEnabled = True
        Me.cmbBlockName.Location = New System.Drawing.Point(15, 27)
        Me.cmbBlockName.Name = "cmbBlockName"
        Me.cmbBlockName.Size = New System.Drawing.Size(357, 21)
        Me.cmbBlockName.TabIndex = 2
        '
        'txtScaleXP
        '
        Me.txtScaleXP.Location = New System.Drawing.Point(12, 73)
        Me.txtScaleXP.Name = "txtScaleXP"
        Me.txtScaleXP.Size = New System.Drawing.Size(150, 20)
        Me.txtScaleXP.TabIndex = 3
        '
        'chkAtivo
        '
        Me.chkAtivo.AutoSize = True
        Me.chkAtivo.Location = New System.Drawing.Point(14, 103)
        Me.chkAtivo.Name = "chkAtivo"
        Me.chkAtivo.Size = New System.Drawing.Size(53, 17)
        Me.chkAtivo.TabIndex = 4
        Me.chkAtivo.Text = "Ativo."
        Me.chkAtivo.UseVisualStyleBackColor = True
        '
        'BtnOK
        '
        Me.BtnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnOK.Location = New System.Drawing.Point(297, 126)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(75, 23)
        Me.BtnOK.TabIndex = 5
        Me.BtnOK.Text = "Aguardando"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(216, 126)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancelar"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'frmRegalComplementAddEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(212, Byte), Integer), CType(CType(212, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(384, 161)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.chkAtivo)
        Me.Controls.Add(Me.txtScaleXP)
        Me.Controls.Add(Me.cmbBlockName)
        Me.Controls.Add(Me.lblScaleXP)
        Me.Controls.Add(Me.lblComplementBlockName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.MaximumSize = New System.Drawing.Size(400, 200)
        Me.MinimumSize = New System.Drawing.Size(400, 200)
        Me.Name = "frmRegalComplementAddEdit"
        Me.Text = "Aguardando..."
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblComplementBlockName As Windows.Forms.Label
    Friend WithEvents lblScaleXP As Windows.Forms.Label
    Friend WithEvents cmbBlockName As Windows.Forms.ComboBox
    Friend WithEvents txtScaleXP As Windows.Forms.TextBox
    Friend WithEvents chkAtivo As Windows.Forms.CheckBox
    Friend WithEvents BtnOK As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
End Class
