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
        If Me.txtAlimBlockName.Text.Trim.Length = 0 Then
            MsgBox("O nome ou parte do nome do bloco alimentador não foi informado.", MsgBoxStyle.Exclamation)
            Me.ActiveControl = Me.txtAlimBlockName
            Exit Sub
        End If
        MyPlugin.RegalBlockNameAlim = Me.txtAlimBlockName.Text.Trim
        Me.Close()
    End Sub

    ''' <summary>
    ''' Carrega os dados do grid
    ''' </summary>
    ''' <param name="conID"></param>
    ''' <param name="cmpID"></param>
    Public Sub LoadData(Optional conID As Object = -1, Optional cmpID As Object = -1)
        Dim DTGeral As System.Data.DataTable
        Dim DTComplemento As System.Data.DataTable
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
                    DTGeral = .CreateConsult
                End With
                .CloseConnection()

                .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL)
                With .SqlSelect
                    .Clear()
                    .AppendLine("SELECT")
                    .AppendLine("RCC.cmpID")
                    .AppendLine(", 'NOME DO BLOCO' = RCC.cmpBlockName")
                    .AppendLine(", 'ESCALA XP' = RCC.cmpScaleXP")
                    .AppendLine(", 'ATIVO' = RCC.ATIVO")
                    .AppendLine("FROM ragaComplementConfig As RCC")
                    .AppendLine("ORDER BY RCC.cmpID")
                    DTComplemento = .CreateConsult
                End With
            End With
        End Using
        Me.dgvComplementaryBlocks.DataSource = DTComplemento
        Me.dgvGeral.DataSource = DTGeral
        Me.dgvGeral.Columns.Item(0).Visible = False
        Me.dgvComplementaryBlocks.Columns.Item(0).Visible = False
        If DTGeral.Rows.Count > 0 Then
            If conID <> -1 Then
                For Index As Integer = 0 To DTGeral.Rows.Count - 1
                    If DTGeral.Rows.Item(Index).Item("conID") = conID Then
                        Me.dgvGeral.Rows.Item(Index).Cells.Item(1).Selected = True
                        Exit For
                    End If
                Next
            Else
                Me.dgvGeral.Rows.Item(0).Cells.Item(1).Selected = True
            End If
        End If
        If DTComplemento.Rows.Count > 0 Then
            If cmpID <> -1 Then
                For Index As Integer = 0 To DTComplemento.Rows.Count - 1
                    If DTComplemento.Rows.Item("cmpID") = cmpID Then
                        Me.dgvComplementaryBlocks.Rows.Item(Index).Cells.Item(1).Selected = True
                    End If
                Next
            End If
        End If
        Me.txtAlimBlockName.Text = MyPlugin.RegalBlockNameAlim
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

    Private Sub btnAddComplementaryBlock_Click(sender As Object, e As EventArgs) Handles btnAddComplementaryBlock.Click
        Dim frmRegalComplementAddEdit As frmRegalComplementAddEdit = Nothing
        Dim RegalComplementAddEditResult As RegalComplementAddEditResult
        Dim DT As System.Data.DataTable
        Dim cmpID As Integer
        Try
            frmRegalComplementAddEdit = New frmRegalComplementAddEdit
            RegalComplementAddEditResult = frmRegalComplementAddEdit.ShowDialog("Criação de Blocos")
            If IsNothing(RegalComplementAddEditResult) = False Then
                Using DB As New adoNetExtension.AdoNetConnect.AdoNet
                    With DB
                        .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL, True)
                        Try
                            With .Fields
                                .ClearAll()
                                .Add("cmpBlockName", RegalComplementAddEditResult.BlockName)
                                .Add("cmpScaleXP", RegalComplementAddEditResult.ScaleXP)
                                .Add("ATIVO", RegalComplementAddEditResult.Ativo)
                                If .AddRegistry("ragaComplementConfig") = True Then
                                    With DB.SqlSelect
                                        .Clear()
                                        .AppendLine("SELECT")
                                        .AppendLine("'cmpID' = Max(RCC.cmpID)")
                                        .AppendLine("FROM")
                                        .AppendLine("ragaComplementConfig AS RCC")
                                        DT = .CreateConsult
                                    End With
                                    cmpID = DT.Rows.Item(0).Item("cmpID")
                                    DB.Commit()
                                    Me.LoadData(cmpID)
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
            If IsNothing(frmRegalComplementAddEdit) = False Then
                frmRegalComplementAddEdit.Dispose()
            End If
        End Try
    End Sub

    Private Sub btnRemoveComplementaryBlock_Click(sender As Object, e As EventArgs) Handles btnRemoveComplementaryBlock.Click
        Dim cmpID As Integer
        Dim cmpBlockName As String
        Try
            If Me.dgvGeral.Rows.Count > 0 Then
                If Me.dgvGeral.SelectedRows.Count > 0 Then
                    cmpID = Me.dgvComplementaryBlocks.SelectedRows.Item(0).Cells.Item("cmpID").Value
                    cmpBlockName = Me.dgvComplementaryBlocks.SelectedRows.Item(0).Cells.Item("NOME DO BLOCO").Value
                    If MsgBox("Deseja excluir o bloco '" & cmpBlockName & "' ?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2) = MsgBoxResult.Yes Then
                        Using DB As New adoNetExtension.AdoNetConnect.AdoNet
                            With DB
                                .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL)
                                With .Fields
                                    .ClearAll()
                                    With .Conditional
                                        .Add("conID", "=", cmpID)
                                    End With
                                    If .DeleteRegistry("ragaComplementConfig", True) = True Then
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

    Private Sub btnEditComplementaryBlock_Click(sender As Object, e As EventArgs) Handles btnEditComplementaryBlock.Click
        Dim frmRegalComplementAddEdit As frmRegalComplementAddEdit = Nothing
        Dim RegalComplementAddEditResult As RegalComplementAddEditResult
        Dim cmpID As Integer
        Dim BlockName As String
        Dim ScaleXP As Double
        Dim ATIVO As Boolean
        Try
            If Me.dgvGeral.Rows.Count > 0 Then
                If Me.dgvGeral.SelectedRows.Count > 0 Then
                    cmpID = Me.dgvGeral.SelectedRows.Item(0).Cells.Item("cmpID").Value
                    BlockName = Me.dgvGeral.SelectedRows.Item(0).Cells.Item("NOME DO BLOCO").Value
                    ScaleXP = Me.dgvGeral.SelectedRows.Item(0).Cells.Item("ESCALA XP").Value
                    ATIVO = Me.dgvGeral.SelectedRows.Item(0).Cells.Item("ATIVO").Value
                    frmRegalComplementAddEdit = New frmRegalComplementAddEdit
                    RegalComplementAddEditResult = New RegalComplementAddEditResult(BlockName, ScaleXP, ATIVO, cmpID)
                    RegalComplementAddEditResult = frmRegalComplementAddEdit.ShowDialog("Edição de blocos", RegalComplementAddEditResult)
                    If IsNothing(RegalComplementAddEditResult) = False Then
                        Using DB As New adoNetExtension.AdoNetConnect.AdoNet
                            With DB
                                .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL)
                                Try
                                    With .Fields
                                        .ClearAll()
                                        .Add("conBlockName", RegalComplementAddEditResult.BlockName)
                                        .Add("conScaleXP", RegalComplementAddEditResult.ScaleXP)
                                        .Add("ATIVO", RegalComplementAddEditResult.Ativo)
                                        With .Conditional
                                            .Add("cmpID", "=", RegalComplementAddEditResult.cmpID)
                                        End With
                                        If .UpdateRegistry("ragaComplementConfig", True) = True Then
                                            Me.LoadData(RegalComplementAddEditResult.cmpID)
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
            If IsNothing(frmRegalComplementAddEdit) = False Then
                frmRegalComplementAddEdit.Dispose()
            End If
        End Try
    End Sub
End Class