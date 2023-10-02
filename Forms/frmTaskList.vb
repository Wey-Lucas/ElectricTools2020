'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2021
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System.Drawing
Imports System.Windows.Forms

Public Class frmTaskList

    'Declarações
    Private _Result As Object
    Private _IsTask As Boolean

    ''' <summary>
    ''' Construtor
    ''' </summary>
    Public Sub New()
        InitializeComponent()
        Me._IsTask = False
    End Sub

    ''' <summary>
    ''' Abre a janela de tarefas
    ''' </summary>
    ''' <param name="Message">Mensagem</param>
    ''' <param name="Tasks">Tarefa</param>
    ''' <returns>Texto</returns>
    Public Overloads Function ShowDialog(Message As String, Tasks As List(Of String), Optional DefaultTask As String = "") As Object
        Dim Task As String
        Dim Button As Button
        Me._IsTask = False
        Me.lblMessage.Text = Message
        With Me.PanelControls.Controls
            Me.PanelControls.AutoScroll = True
            Tasks.Reverse()
            For Index As Integer = 0 To Tasks.Count - 1
                Task = Tasks.Item(Index)
                Button = New Button
                With Button
                    .Name = "Button" & Index.ToString
                    .Text = Task
                    If Task = DefaultTask Then
                        .BackColor = Color.Lavender
                        .ForeColor = Color.DarkRed
                    Else
                        .BackColor = Me.BackColor
                        .ForeColor = Color.MediumBlue
                    End If
                    .Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    .TextAlign = Drawing.ContentAlignment.MiddleCenter
                    .FlatStyle = FlatStyle.Flat
                    .FlatAppearance.BorderSize = 0
                    .FlatAppearance.MouseDownBackColor = Color.Transparent
                    .Dock = DockStyle.Top
                    .MinimumSize = New Size(Me.PanelControls.Width - 1, 45)
                    .Image = My.Resources.TaskLeader
                    .ImageAlign = ContentAlignment.MiddleLeft
                    .Cursor = System.Windows.Forms.Cursors.Hand
                    .Show()
                End With
                Me.ToolTip.SetToolTip(Button, "Acione esta opção para aplicar " & Task & ".")
                .Add(Button)
                If DefaultTask <> "" Then
                    If Button.Text = DefaultTask Then
                        Me.AcceptButton = Button
                        Me.ActiveControl = Button
                    End If
                Else
                    Me.AcceptButton = Me.cmdCancel
                    Me.ActiveControl = Me.cmdCancel
                End If
                If DefaultTask <> "" Then
                    Me.lblHelp.Text = "Tecle [ENTER] para " & DefaultTask & "."
                Else
                    Me.lblHelp.Text = "Tecle [ENTER] para cancelar."
                End If
                AddHandler Button.Click, AddressOf TaskClick
            Next
        End With
        Me.ShowDialog()
        Return Me._Result
    End Function

    ''' <summary>
    ''' Abre a janela de tarefas
    ''' </summary>
    ''' <param name="Message">Mensagem</param>
    ''' <param name="Tasks">Tarefa</param>
    ''' <param name="DefaultTask">Tarefa padrão</param>
    ''' <param name="ImageList">Lista de imagens</param>
    ''' <returns>Tarefa (Task)</returns>
    Public Overloads Function ShowDialog(Message As String, Tasks As List(Of Task), Optional DefaultTask As Task = Nothing, Optional ImageList As ImageList = Nothing) As Object
        Dim Task As Task
        Dim Button As Button
        Me._IsTask = True
        Me.lblMessage.Text = Message
        With Me.PanelControls.Controls
            Me.PanelControls.AutoScroll = True
            Tasks.Reverse()
            For Index As Integer = 0 To Tasks.Count - 1
                Task = Tasks.Item(Index)
                Button = New Button
                With Button
                    .Name = "Button" & Index.ToString
                    .Text = Task.Value
                    .Tag = Task
                    If Task.Tag.Equals(DefaultTask.Tag) = True Then
                        .BackColor = Color.Lavender
                        .ForeColor = Color.DarkRed
                    Else
                        .BackColor = Me.BackColor
                        .ForeColor = Color.MediumBlue
                    End If
                    .Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    .TextAlign = Drawing.ContentAlignment.MiddleCenter
                    .FlatStyle = FlatStyle.Flat
                    .FlatAppearance.BorderSize = 0
                    .FlatAppearance.MouseDownBackColor = Color.Transparent
                    .Dock = DockStyle.Top
                    .MinimumSize = New Size(Me.PanelControls.Width - 1, 45)
                    If IsNothing(ImageList) = False Then
                        .ImageList = ImageList
                        .ImageIndex = Task.ImageIndex
                    Else
                        .Image = My.Resources.TaskLeader
                    End If
                    .ImageAlign = ContentAlignment.MiddleLeft
                    .Cursor = System.Windows.Forms.Cursors.Hand
                    .Show()
                End With
                Me.ToolTip.SetToolTip(Button, "Acione esta opção para aplicar " & Task.Value & ".")
                .Add(Button)
                If IsNothing(DefaultTask) = False Then
                    If Button.Tag.Equals(DefaultTask) = True Then
                        Me.AcceptButton = Button
                        Me.ActiveControl = Button
                    End If
                Else
                    Me.AcceptButton = Me.cmdCancel
                    Me.ActiveControl = Me.cmdCancel
                End If
                If IsNothing(DefaultTask) = False Then
                    Me.lblHelp.Text = "Tecle [ENTER] para " & DefaultTask.Value & "."
                Else
                    Me.lblHelp.Text = "Tecle [ENTER] para cancelar."
                End If
                AddHandler Button.Click, AddressOf TaskClick
            Next
        End With
        Me.ShowDialog()
        Return Me._Result
    End Function

    ''' <summary>
    ''' Detecta a seleção de uma tarefa
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TaskClick(sender As Object, e As EventArgs)
        If Me._IsTask = True Then
            Me._Result = sender.tag
        Else
            Me._Result = sender.text
        End If
        Me.Close()
    End Sub

    ''' <summary>
    ''' Cancelar
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me._Result = Nothing
        Me.Close()
    End Sub

End Class

''' <summary>
''' Retorno de frmTaskList
''' </summary>
Public Class Task

    'Declarações
    Private _Value As String
    Private _ImageIndex As Integer
    Private _Tag As Object

    ''' <summary>
    ''' Construtor
    ''' </summary>
    ''' <param name="Value">Valor de exibição</param>
    ''' <param name="Tag">Valor do usuário</param>
    ''' <param name="ImageIndex">Posição da imagem a ser exibida</param>
    Public Sub New(Value As String, Tag As Object, Optional ImageIndex As Integer = -1)
        Me._Value = Value
        Me._Tag = Tag
        Me._ImageIndex = ImageIndex
    End Sub

    ''' <summary>
    ''' Valor de exibição
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Value As String
        Get
            Return Me._Value
        End Get
    End Property

    ''' <summary>
    ''' Posição da imagem a ser exibida
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ImageIndex As Integer
        Get
            Return Me._ImageIndex
        End Get
    End Property

    ''' <summary>
    ''' Valor do usuário
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Tag As Object
        Get
            Return Me._Tag
        End Get
    End Property

End Class