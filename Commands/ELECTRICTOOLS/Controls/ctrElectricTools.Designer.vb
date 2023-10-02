<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ctrElectricTools
    Inherits System.Windows.Forms.UserControl

    'O UserControl substitui o descarte para limpar a lista de componentes.
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
        Me.Panel = New System.Windows.Forms.Panel()
        Me.gbAlimentadores = New System.Windows.Forms.GroupBox()
        Me.cmdREGAL = New System.Windows.Forms.Button()
        Me.msElectricTools = New System.Windows.Forms.MenuStrip()
        Me.tsmTools = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmRegalConfig = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmTopics = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmConfig = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel.SuspendLayout()
        Me.gbAlimentadores.SuspendLayout()
        Me.msElectricTools.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel
        '
        Me.Panel.BackColor = System.Drawing.Color.FromArgb(CType(CType(212, Byte), Integer), CType(CType(212, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel.Controls.Add(Me.gbAlimentadores)
        Me.Panel.Controls.Add(Me.msElectricTools)
        Me.Panel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel.Location = New System.Drawing.Point(0, 0)
        Me.Panel.Name = "Panel"
        Me.Panel.Size = New System.Drawing.Size(250, 452)
        Me.Panel.TabIndex = 0
        '
        'gbAlimentadores
        '
        Me.gbAlimentadores.Controls.Add(Me.cmdREGAL)
        Me.gbAlimentadores.ForeColor = System.Drawing.Color.Black
        Me.gbAlimentadores.Location = New System.Drawing.Point(5, 27)
        Me.gbAlimentadores.Name = "gbAlimentadores"
        Me.gbAlimentadores.Size = New System.Drawing.Size(234, 54)
        Me.gbAlimentadores.TabIndex = 3
        Me.gbAlimentadores.TabStop = False
        Me.gbAlimentadores.Text = "Alimentadores"
        '
        'cmdREGAL
        '
        Me.cmdREGAL.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdREGAL.ForeColor = System.Drawing.Color.Black
        Me.cmdREGAL.Image = Global.ElectricTools2020.My.Resources.Resources.PNG_EDIT
        Me.cmdREGAL.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdREGAL.Location = New System.Drawing.Point(6, 19)
        Me.cmdREGAL.Name = "cmdREGAL"
        Me.cmdREGAL.Size = New System.Drawing.Size(75, 23)
        Me.cmdREGAL.TabIndex = 2
        Me.cmdREGAL.Text = "Registro"
        Me.cmdREGAL.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdREGAL.UseVisualStyleBackColor = True
        '
        'msElectricTools
        '
        Me.msElectricTools.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.msElectricTools.GripStyle = System.Windows.Forms.ToolStripGripStyle.Visible
        Me.msElectricTools.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmTools, Me.tsmHelp, Me.tsmConfig})
        Me.msElectricTools.Location = New System.Drawing.Point(0, 0)
        Me.msElectricTools.Name = "msElectricTools"
        Me.msElectricTools.Size = New System.Drawing.Size(248, 24)
        Me.msElectricTools.TabIndex = 1
        '
        'tsmTools
        '
        Me.tsmTools.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmRegalConfig})
        Me.tsmTools.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.tsmTools.ImageTransparentColor = System.Drawing.Color.White
        Me.tsmTools.Name = "tsmTools"
        Me.tsmTools.Size = New System.Drawing.Size(84, 20)
        Me.tsmTools.Text = "Ferramentas"
        '
        'tsmRegalConfig
        '
        Me.tsmRegalConfig.Name = "tsmRegalConfig"
        Me.tsmRegalConfig.Size = New System.Drawing.Size(211, 22)
        Me.tsmRegalConfig.Text = "Registro de alimentadores"
        '
        'tsmHelp
        '
        Me.tsmHelp.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsmHelp.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.tsmHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmTopics, Me.tsmAbout})
        Me.tsmHelp.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.tsmHelp.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.tsmHelp.ImageTransparentColor = System.Drawing.Color.White
        Me.tsmHelp.Name = "tsmHelp"
        Me.tsmHelp.Size = New System.Drawing.Size(50, 20)
        Me.tsmHelp.Text = "Ajuda"
        '
        'tsmTopics
        '
        Me.tsmTopics.Name = "tsmTopics"
        Me.tsmTopics.Size = New System.Drawing.Size(162, 22)
        Me.tsmTopics.Text = "Tópicos de ajuda"
        '
        'tsmAbout
        '
        Me.tsmAbout.Name = "tsmAbout"
        Me.tsmAbout.Size = New System.Drawing.Size(162, 22)
        Me.tsmAbout.Text = "Autoria e versão"
        '
        'tsmConfig
        '
        Me.tsmConfig.Name = "tsmConfig"
        Me.tsmConfig.Size = New System.Drawing.Size(96, 20)
        Me.tsmConfig.Text = "Configurações"
        '
        'ctrElectricTools
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(212, Byte), Integer), CType(CType(212, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.Controls.Add(Me.Panel)
        Me.Name = "ctrElectricTools"
        Me.Size = New System.Drawing.Size(250, 452)
        Me.Panel.ResumeLayout(False)
        Me.Panel.PerformLayout()
        Me.gbAlimentadores.ResumeLayout(False)
        Me.msElectricTools.ResumeLayout(False)
        Me.msElectricTools.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel As Windows.Forms.Panel
    Friend WithEvents msElectricTools As Windows.Forms.MenuStrip
    Friend WithEvents tsmTools As Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmHelp As Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmTopics As Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmAbout As Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmConfig As Windows.Forms.ToolStripMenuItem
    Private WithEvents cmdREGAL As Windows.Forms.Button
    Friend WithEvents gbAlimentadores As Windows.Forms.GroupBox
    Friend WithEvents tsmRegalConfig As Windows.Forms.ToolStripMenuItem
End Class
