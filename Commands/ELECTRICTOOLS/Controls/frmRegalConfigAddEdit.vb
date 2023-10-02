Imports System.Windows.Controls
Imports System.Windows.Forms
Imports ElectricTools2020.ElectricTools

''' <summary>
''' Permite adicionar e editar
''' </summary>
Public Class frmRegalConfigAddEdit

    'Declarações
    Private _Result As RegalConfigAddEditResult

    ''' <summary>
    ''' Abre para adicionar e editar
    ''' </summary>
    ''' <param name="Title"></param>
    ''' <param name="RegalConfigAddEditResult"></param>
    ''' <returns></returns>
    Public Overloads Function ShowDialog(Title As String, Optional RegalConfigAddEditResult As RegalConfigAddEditResult = Nothing) As RegalConfigAddEditResult
        Me.Text = Title
        Me._Result = RegalConfigAddEditResult
        If IsNothing(RegalConfigAddEditResult) = True Then
            Me.cmdOK.Text = "Adicionar"
            Me.cmbBlockName.Text = ""
            Me.txtScaleXP.Text = "0.50"
            Me.txtLines.Text = "10"
            Me.chkAtivo.Checked = True
        Else
            Me.cmdOK.Text = "Editar"
            Me.cmbBlockName.Text = RegalConfigAddEditResult.BlockName
            Me.txtScaleXP.Text = RegalConfigAddEditResult.ScaleXP.ToString
            Me.txtLines.Text = RegalConfigAddEditResult.Lines.ToString
            Me.chkAtivo.Checked = RegalConfigAddEditResult.Ativo
        End If
        Me.ShowDialog()
        Return Me._Result
    End Function

    ''' <summary>
    ''' Cancel
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me._Result = Nothing
        Me.Close()
    End Sub

    ''' <summary>
    ''' OK
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        'Declarações
        Dim DT As System.Data.DataTable

        'Validações
        If Me.cmbBlockName.Text.Trim.Length = 0 Then
            MsgBox("O nome do bloco não foi informado.", MsgBoxStyle.Exclamation)
            Me.ActiveControl = Me.cmbBlockName
            Exit Sub
        End If
        If Me.txtScaleXP.Text.Trim.Length = 0 Then
            MsgBox("A escala não foi informada.", MsgBoxStyle.Exclamation)
            Me.ActiveControl = Me.txtScaleXP
            Exit Sub
        ElseIf CDbl(Me.txtScaleXP.Text) <= 0 Then
            MsgBox("A escala não pode ser menor ou igual a zero.", MsgBoxStyle.Exclamation)
            Me.ActiveControl = Me.txtScaleXP
            Exit Sub
        End If
        If Me.txtLines.Text.Trim.Length = 0 Then
            MsgBox("O número de linhas não foi informado.", MsgBoxStyle.Exclamation)
            Me.ActiveControl = Me.txtLines
            Exit Sub
        ElseIf CInt(txtLines.Text) < 1 Then
            MsgBox("O número de linhas não pode ser menor que um.", MsgBoxStyle.Exclamation)
            Me.ActiveControl = Me.txtLines
            Exit Sub
        End If

        Using DB As New adoNetExtension.AdoNetConnect.AdoNet
            With DB
                .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL)
                With .SqlSelect
                    .Clear()
                    .AppendLine("SELECT")
                    .AppendLine("CON.conID")
                    .AppendLine(", CON.conBlockName")
                    .AppendLine("FROM")
                    .AppendLine("ragaConfig AS CON")
                    .AppendLine("WHERE")
                    .AppendLine("(")
                    .AppendLine("CON.conBlockName = '" & Me.cmbBlockName.Text.Trim & "'")
                    If IsNothing(Me._Result) = False Then
                        .AppendLine("AND")
                        .AppendLine("CON.conID <> " & Me._Result.conID)
                    End If
                    .AppendLine(")")
                    DT = .CreateConsult
                End With
                .CloseConnection()
            End With
        End Using
        If DT.Rows.Count > 0 Then
            MsgBox("O nome do bloco já foi cadastrado.", MsgBoxStyle.Exclamation)
            Me.ActiveControl = Me.cmbBlockName
            Exit Sub
        End If

        'Retorno
        Me._Result = New RegalConfigAddEditResult(Me.cmbBlockName.Text, Me.txtScaleXP.Text, Me.txtLines.Text, Me.chkAtivo.Checked, If(IsNothing(Me._Result) = False, Me._Result.conID, -1))

        'Fecha
        Me.Close()

    End Sub

    ''' <summary>
    ''' Escala
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub lblScaleXP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles lblScaleXP.KeyPress
        e.Handled = Engine2.KeyPressMonitor.KeyPress(Engine2.KeyPressMonitor.eInputType._Real, sender, e, 3)
    End Sub

    ''' <summary>
    ''' Linhas
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub lblLines_KeyPress(sender As Object, e As KeyPressEventArgs) Handles lblLines.KeyPress
        e.Handled = Engine2.KeyPressMonitor.KeyPress(Engine2.KeyPressMonitor.eInputType._Integer, sender, e)
    End Sub

    ''' <summary>
    ''' Carrega os nomes dos blocos
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmbBlockName_DropDown(sender As Object, e As EventArgs) Handles cmbBlockName.DropDown
        Dim DT As DataTable
        Using DB As New adoNetExtension.AdoNetConnect.AdoNet
            With DB
                .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL)
                With .SqlSelect
                    .Clear()
                    .AppendLine("SELECT")
                    .AppendLine("CON.conBlockName")
                    .AppendLine("FROM")
                    .AppendLine("ragaConfig AS CON")
                    DT = .CreateConsult
                End With
                .CloseConnection()
            End With
        End Using
        With Me.cmbBlockName
            .Items.Clear()
            For Each DR As DataRow In DT.Rows
                .Items.Add(DR.Item("conBlockName"))
            Next
            .Sorted = True
        End With
    End Sub

End Class

''' <summary>
''' Retorno de frmRegalConfigAddEdit
''' </summary>
Public Class RegalConfigAddEditResult

    'Declarações
    Private _conID As Integer
    Private _BlockName As String
    Private _ScaleXP As Double
    Private _Lines As Integer
    Private _Ativo As Boolean

    ''' <summary>
    ''' Construtor
    ''' </summary>
    ''' <param name="BlockName">Nome do bloco</param>
    ''' <param name="ScaleXP">Escala XP</param>
    ''' <param name="Lines">Número de linhas</param>
    ''' <param name="Ativo">Determina se o registro é Ativo</param>
    Public Sub New(BlockName As String, ScaleXP As Double, Lines As Integer, Ativo As Boolean, Optional conID As Integer = -1)
        Me._BlockName = BlockName.Trim
        Me._ScaleXP = ScaleXP
        Me._Lines = Lines
        Me._conID = conID
        Me._Ativo = Ativo
    End Sub

    ''' <summary>
    ''' Nome do bloco
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property BlockName As String
        Get
            Return Me._BlockName
        End Get
    End Property

    ''' <summary>
    ''' Escala XP
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ScaleXP As Double
        Get
            Return Me._ScaleXP
        End Get
    End Property

    ''' <summary>
    ''' Número de linhas
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Lines As Integer
        Get
            Return Me._Lines
        End Get
    End Property

    ''' <summary>
    ''' ID
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property conID As Integer
        Get
            Return Me._conID
        End Get
    End Property

    ''' <summary>
    ''' Determina se o registro é Ativo
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Ativo As Boolean
        Get
            Return Me._Ativo
        End Get
    End Property

End Class