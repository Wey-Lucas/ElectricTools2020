Imports System.Windows.Controls
Imports System.Windows.Forms
Imports ElectricTools2020.ElectricTools

Public Class frmRegalComplementAddEdit

    Private _Result As RegalComplementAddEditResult
    Public Overloads Function ShowDialog(Title As String, Optional RegalComplementAddEditResult As RegalComplementAddEditResult = Nothing) As RegalComplementAddEditResult
        Me.Text = Title
        Me._Result = RegalComplementAddEditResult
        If IsNothing(RegalComplementAddEditResult) = True Then
            Me.BtnOK.Text = "Adicionar"
            Me.cmbBlockName.Text = ""
            Me.txtScaleXP.Text = "0.05"
            Me.chkAtivo.Checked = True
        Else
            Me.BtnOK.Text = "Editar"
            Me.cmbBlockName.Text = RegalComplementAddEditResult.BlockName
            Me.txtScaleXP.Text = RegalComplementAddEditResult.ScaleXP
            Me.chkAtivo.Checked = RegalComplementAddEditResult.Ativo
        End If
        Me.ShowDialog()
        Return Me._Result
    End Function

    ''' <summary>
    ''' Alimenta a ComboBox com os nomes dos blocos
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmbBlockName_DropDown(sender As Object, e As EventArgs) Handles cmbBlockName.DropDown
        Dim DT As DataTable
        Using DB As New adoNetExtension.AdoNetConnect.AdoNet
            With DB
                .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL)
                With .SqlSelect
                    .AppendLine("SELECT")
                    .AppendLine("RCC.cmpBlockName")
                    .AppendLine("FROM ragaComplementConfig AS RCC")
                    DT = .CreateConsult
                End With
                .CloseConnection()
            End With
        End Using
        With Me.cmbBlockName
            .Items.Clear()
            For Each DR As DataRow In DT.Rows
                .Items.Add(DR.Item("cmpBlockName"))
            Next
            .Sorted = True
        End With
    End Sub

    ''' <summary>
    ''' Trataiva de valores inseridos na TextBox de Escala
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub txtScaleXP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtScaleXP.KeyPress
        e.Handled = Engine2.KeyPressMonitor.KeyPress(Engine2.KeyPressMonitor.eInputType._Real, sender, e, 3)
    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs) Handles BtnOK.Click

        Dim DT As System.Data.DataTable

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

        Using DB As New adoNetExtension.AdoNetConnect.AdoNet
            With DB
                .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL)
                With .SqlSelect
                    .Clear()
                    .AppendLine("SELECT")
                    .AppendLine("RCC.cmpID")
                    .AppendLine(", RCC.cmpBlockName")
                    .AppendLine("FROM ragaComplementConfig AS RCC")
                    .AppendLine("WHERE")
                    .AppendLine("(")
                    .AppendLine("RCC.cmpBlockName = '" & Me.cmbBlockName.Text.Trim & "'")
                    If IsNothing(Me._Result) = False Then
                        .AppendLine("AND RCC.cmpID <> " & Me._Result.cmpID)
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

        Me._Result = New RegalComplementAddEditResult(Me.cmbBlockName.Text, Me.txtScaleXP.Text, Me.chkAtivo.Checked, If(IsNothing(Me._Result) = False, Me._Result.cmpID, -1))

        Me.Close()

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me._Result = Nothing
        Me.Close()
    End Sub
End Class

''' <summary>
''' Retorno de frmRegalComplementAddEdit
''' </summary>
Public Class RegalComplementAddEditResult

    'Declarações
    Private _cmpID As Integer
    Private _BlockName As String
    Private _ScaleXP As Double
    Private _Ativo As Boolean

    ''' <summary>
    ''' Construtor
    ''' </summary>
    ''' <param name="BlockName"></param>
    ''' <param name="ScaleXP"></param>
    ''' <param name="Ativo"></param>
    ''' <param name="cmpID"></param>
    Public Sub New(BlockName As String, ScaleXP As Double, Ativo As Boolean, Optional cmpID As Integer = -1)
        Me._BlockName = BlockName
        Me._ScaleXP = ScaleXP
        Me._Ativo = Ativo
        Me._cmpID = cmpID
    End Sub

    Public ReadOnly Property BlockName As String
        Get
            Return Me._BlockName
        End Get
    End Property

    Public ReadOnly Property ScaleXP As Double
        Get
            Return Me._ScaleXP
        End Get
    End Property

    Public ReadOnly Property Ativo As Boolean
        Get
            Return Me._Ativo
        End Get
    End Property

    Public ReadOnly Property cmpID As Integer
        Get
            Return Me._cmpID
        End Get
    End Property

End Class