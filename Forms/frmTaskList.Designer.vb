<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmTaskList
    Inherits System.Windows.Forms.Form

    'Descartar substituições de formulário para limpar a lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.PanelButton = New System.Windows.Forms.Panel()
        Me.lblHelp = New System.Windows.Forms.Label()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.PanelControls = New System.Windows.Forms.Panel()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.pnGeral = New System.Windows.Forms.Panel()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.PanelButton.SuspendLayout()
        Me.pnGeral.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelButton
        '
        Me.PanelButton.BackColor = System.Drawing.SystemColors.Control
        Me.PanelButton.Controls.Add(Me.lblHelp)
        Me.PanelButton.Controls.Add(Me.cmdCancel)
        Me.PanelButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelButton.Location = New System.Drawing.Point(0, 308)
        Me.PanelButton.Name = "PanelButton"
        Me.PanelButton.Size = New System.Drawing.Size(298, 40)
        Me.PanelButton.TabIndex = 2
        '
        'lblHelp
        '
        Me.lblHelp.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblHelp.ForeColor = System.Drawing.Color.DarkGreen
        Me.lblHelp.Location = New System.Drawing.Point(11, 9)
        Me.lblHelp.Name = "lblHelp"
        Me.lblHelp.Size = New System.Drawing.Size(180, 23)
        Me.lblHelp.TabIndex = 1
        Me.lblHelp.Text = "Help..."
        Me.lblHelp.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Crimson
        Me.cmdCancel.Location = New System.Drawing.Point(197, 5)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(95, 30)
        Me.cmdCancel.TabIndex = 0
        Me.cmdCancel.Text = "Cancelar"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'PanelControls
        '
        Me.PanelControls.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PanelControls.Location = New System.Drawing.Point(30, 37)
        Me.PanelControls.Name = "PanelControls"
        Me.PanelControls.Size = New System.Drawing.Size(234, 265)
        Me.PanelControls.TabIndex = 1
        '
        'lblMessage
        '
        Me.lblMessage.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessage.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblMessage.Image = My.Resources.ALERTMINI
        Me.lblMessage.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblMessage.Location = New System.Drawing.Point(0, 0)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(298, 34)
        Me.lblMessage.TabIndex = 0
        Me.lblMessage.Text = "Information"
        Me.lblMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnGeral
        '
        Me.pnGeral.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnGeral.Controls.Add(Me.lblMessage)
        Me.pnGeral.Controls.Add(Me.PanelControls)
        Me.pnGeral.Controls.Add(Me.PanelButton)
        Me.pnGeral.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnGeral.Location = New System.Drawing.Point(0, 0)
        Me.pnGeral.Name = "pnGeral"
        Me.pnGeral.Size = New System.Drawing.Size(300, 350)
        Me.pnGeral.TabIndex = 3
        '
        'frmTaskList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(300, 350)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnGeral)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmTaskList"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmTaskList"
        Me.PanelButton.ResumeLayout(False)
        Me.pnGeral.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PanelButton As Windows.Forms.Panel
    Friend WithEvents cmdCancel As Windows.Forms.Button
    Friend WithEvents PanelControls As Windows.Forms.Panel
    Friend WithEvents lblMessage As Windows.Forms.Label
    Friend WithEvents pnGeral As Windows.Forms.Panel
    Friend WithEvents lblHelp As Windows.Forms.Label
    Friend WithEvents ToolTip As Windows.Forms.ToolTip
End Class
