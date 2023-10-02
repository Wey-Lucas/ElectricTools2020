Imports System.Windows.Forms
Imports ElectricTools2020.ElectricTools

Public Class frmRegalConfig

    ''' <summary>
    ''' Construtor
    ''' </summary>
    Public Sub New()
        InitializeComponent()
        Me.LoadData()
    End Sub

    ''' <summary>
    ''' Cancelar
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' Carrega os dados do grid
    ''' </summary>
    ''' <param name="conID"></param>
    Public Sub LoadData(Optional conID As Object = -1)
        Dim DT As System.Data.DataTable
        Using DB As New adoNetExtension.AdoNetConnect.AdoNet
            With DB
                .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL)
                With .SqlSelect
                    .Clear()
                    .AppendLine("SELECT")
                    .AppendLine("CON.conID")
                    .AppendLine(", 'NOME DO BLOCO' = CON.conBlockName")
                    .AppendLine(", 'ESCALA XP' = CON.conScaleXP")
                    .AppendLine(", 'NÚMERO DE LINHAS' = CON.conLines")
                    .AppendLine(", 'ATIVO' = CON.ATIVO")
                    .AppendLine("FROM")
                    .AppendLine("ragaConfig AS CON")
                    .AppendLine("ORDER BY CON.conBlockName")
                    DT = .CreateConsult
                End With
                .CloseConnection()
            End With
        End Using
        Me.dgvGeral.DataSource = DT
        Me.dgvGeral.Columns.Item(0).Visible = False
        If DT.Rows.Count > 0 Then
            If conID <> -1 Then
                For Index As Integer = 0 To DT.Rows.Count - 1
                    If DT.Rows.Item(Index).Item("conID") = conID Then
                        Me.dgvGeral.Rows.Item(Index).Cells.Item(1).Selected = True
                        Exit For
                    End If
                Next
            Else
                Me.dgvGeral.Rows.Item(0).Cells.Item(1).Selected = True
            End If
        End If
    End Sub

    ''' <summary>
    ''' Adicionar
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdAdicionar_Click(sender As Object, e As EventArgs) Handles cmdAdicionar.Click
        Dim frmRegalConfigAddEdit As frmRegalConfigAddEdit = Nothing
        Dim RegalConfigAddEditResult As RegalConfigAddEditResult
        Dim DT As System.Data.DataTable
        Dim conID As Integer
        Try
            frmRegalConfigAddEdit = New frmRegalConfigAddEdit
            RegalConfigAddEditResult = frmRegalConfigAddEdit.ShowDialog("Criação de blocos")
            If IsNothing(RegalConfigAddEditResult) = False Then
                Using DB As New adoNetExtension.AdoNetConnect.AdoNet
                    With DB
                        .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL, True)
                        Try
                            With .Fields
                                .ClearAll()
                                .Add("conBlockName", RegalConfigAddEditResult.BlockName)
                                .Add("conScaleXP", RegalConfigAddEditResult.ScaleXP)
                                .Add("conLines", RegalConfigAddEditResult.Lines)
                                .Add("ATIVO", RegalConfigAddEditResult.Ativo)
                                If .AddRegistry("ragaConfig") = True Then
                                    With DB.SqlSelect
                                        .Clear()
                                        .AppendLine("SELECT")
                                        .AppendLine("'conID' = Max(CON.conID)")
                                        .AppendLine("FROM")
                                        .AppendLine("ragaConfig AS CON")
                                        DT = .CreateConsult
                                    End With
                                    conID = DT.Rows.Item(0).Item("conID")
                                    DB.Commit()
                                    Me.LoadData(conID)
                                Else
                                    MsgBox("Impossível criar o registro.", MsgBoxStyle.Exclamation)
                                End If
                            End With
                        Catch
                            DB.Rollback()
                        End Try
                        .CloseConnection()
                    End With
                End Using
            End If
        Catch ex As System.Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        Finally
            If IsNothing(frmRegalConfigAddEdit) = False Then
                frmRegalConfigAddEdit.Dispose()
            End If
        End Try
    End Sub



    ''' <summary>
    ''' Excluir
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdRemove_Click(sender As Object, e As EventArgs) Handles cmdRemove.Click
        Dim conID As Integer
        Dim conBlockName As String
        Try
            If Me.dgvGeral.Rows.Count > 0 Then
                If Me.dgvGeral.SelectedRows.Count > 0 Then
                    conID = Me.dgvGeral.SelectedRows.Item(0).Cells.Item("conID").Value
                    conBlockName = Me.dgvGeral.SelectedRows.Item(0).Cells.Item("NOME DO BLOCO").Value
                    If MsgBox("Deseja excluir o bloco '" & conBlockName & "' ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2) = MsgBoxResult.Yes Then
                        Using DB As New adoNetExtension.AdoNetConnect.AdoNet
                            With DB
                                .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL)
                                With .Fields
                                    .ClearAll()
                                    With .Conditional
                                        .Add("conID", "=", conID)
                                    End With
                                    If .DeleteRegistry("ragaConfig", True) = True Then
                                        Me.LoadData()
                                    Else
                                        MsgBox("Impossível excluir o registro.", MsgBoxStyle.Exclamation)
                                    End If
                                End With
                                .CloseConnection()
                            End With
                        End Using
                    End If
                Else
                    MsgBox("Nenhum registro foi selecionado para exclusão.", MsgBoxStyle.Exclamation)
                End If
            Else
                MsgBox("Nenhum registro foi localizado para exclusão.", MsgBoxStyle.Exclamation)
            End If
        Catch ex As System.Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    ''' <summary>
    ''' Editar
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        Dim frmRegalConfigAddEdit As frmRegalConfigAddEdit = Nothing
        Dim RegalConfigAddEditResult As RegalConfigAddEditResult
        Dim conID As Integer
        Dim BlockName As String
        Dim ScaleXP As Double
        Dim Lines As Integer
        Dim ATIVO As Boolean
        Try
            If Me.dgvGeral.Rows.Count > 0 Then
                If Me.dgvGeral.SelectedRows.Count > 0 Then
                    conID = Me.dgvGeral.SelectedRows.Item(0).Cells.Item("conID").Value
                    BlockName = Me.dgvGeral.SelectedRows.Item(0).Cells.Item("NOME DO BLOCO").Value
                    ScaleXP = Me.dgvGeral.SelectedRows.Item(0).Cells.Item("ESCALA XP").Value
                    Lines = Me.dgvGeral.SelectedRows.Item(0).Cells.Item("NÚMERO DE LINHAS").Value
                    ATIVO = Me.dgvGeral.SelectedRows.Item(0).Cells.Item("ATIVO").Value
                    frmRegalConfigAddEdit = New frmRegalConfigAddEdit
                    RegalConfigAddEditResult = New RegalConfigAddEditResult(BlockName, ScaleXP, Lines, ATIVO, conID)
                    RegalConfigAddEditResult = frmRegalConfigAddEdit.ShowDialog("Edição de blocos", RegalConfigAddEditResult)
                    If IsNothing(RegalConfigAddEditResult) = False Then
                        Using DB As New adoNetExtension.AdoNetConnect.AdoNet
                            With DB
                                .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL)
                                Try
                                    With .Fields
                                        .ClearAll()
                                        .Add("conBlockName", RegalConfigAddEditResult.BlockName)
                                        .Add("conScaleXP", RegalConfigAddEditResult.ScaleXP)
                                        .Add("conLines", RegalConfigAddEditResult.Lines)
                                        .Add("ATIVO", RegalConfigAddEditResult.Ativo)
                                        With .Conditional
                                            .Add("conID", "=", RegalConfigAddEditResult.conID)
                                        End With
                                        If .UpdateRegistry("ragaConfig", True) = True Then
                                            Me.LoadData(RegalConfigAddEditResult.conID)
                                        Else
                                            MsgBox("Impossível editar o registro.", MsgBoxStyle.Exclamation)
                                        End If
                                    End With
                                Catch
                                    DB.Rollback()
                                End Try
                                .CloseConnection()
                            End With
                        End Using
                    End If
                Else
                    MsgBox("Nenhum registro foi selecionado para edição.", MsgBoxStyle.Exclamation)
                End If
            Else
                MsgBox("Nenhum registro foi localizado para edição.", MsgBoxStyle.Exclamation)
            End If
        Catch ex As System.Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        Finally
            If IsNothing(frmRegalConfigAddEdit) = False Then
                frmRegalConfigAddEdit.Dispose()
            End If
        End Try
    End Sub
    Private Sub dgvGeral_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvGeral.CellDoubleClick
        If e.RowIndex <> -1 And e.ColumnIndex <> -1 Then
            Me.cmdEdit_Click(Nothing, Nothing)
        End If
    End Sub

End Class