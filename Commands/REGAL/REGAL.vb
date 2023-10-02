'=========================================================================================================='
'DESENVOLVIDO POR: LUCAS LEÔNCIO WEY DA SILVA (FORMANDO EM ANALISTA DE SISTEMAS)
'EM:2023
'=========================================================================================================='
'CONTATO: wey.lucas1@gmail.com \ (011) 99940-2202
'=========================================================================================================='

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows
Imports ElectricTools2020.ElectricTools
Imports ElectricTools2020.Engine2
Imports SPHEUpdate47
Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Linq

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Commands.REGAL))>
Namespace Commands
    Public Class REGAL

        'Declarações
        Private Shared WithEvents _PaletteSetControl As Engine2.MyPaletteSet
        Private Shared _CurrentPalette As Palette
        Private Shared _ctrElectricTools As ctrElectricTools
        Private Shared WithEvents _SelectEntityInBlockPlane As SelectEntityInBlockPlane
        Private Shared _DtConfig As System.Data.DataTable

        ''' <summary>
        ''' Retorna ctrDTSIMP
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property ctrElectricTools As ctrElectricTools
            Get
                Return _ctrElectricTools
            End Get
        End Property

        ''' <summary>
        ''' Filtra os blocos válidos
        ''' </summary>
        ''' <param name="Entity"></param>
        ''' <param name="IsValid"></param>
        ''' <param name="CurrentTransaction"></param>
        Private Shared Sub _SelectEntityInBlockPlane_SelectionFilter(Entity As Entity, ByRef IsValid As Boolean, CurrentTransaction As Transaction) Handles _SelectEntityInBlockPlane.SelectionFilter
            Dim DR() As DataRow
            Dim BlockReference As BlockReference
            Select Case Entity.GetType
                Case GetType(BlockReference)
                    BlockReference = Entity
                    DR = _DtConfig.Select("conBlockName = '" & BlockReference.Name & "'")
                    If DR.Length > 0 Then
                        IsValid = True
                    Else
                        IsValid = False
                    End If
                Case Else
                    IsValid = False
            End Select
        End Sub

        ''' <summary>
        ''' REGAL
        ''' </summary>
        <CommandMethod("REGAL", CommandFlags.Modal)>
        Public Shared Sub REGAL()

            If Support.Lic.GetLic = False Then
                Exit Sub
            End If

            Support.Lic.Add(Support.Lic.eArchitecture.AutoCAD, My.Application.Info.AssemblyName, "ELECTRICALTOOLS")

            'Declarações
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim BlocksMode As New List(Of BlockReference)
            Dim BlocksPlanta As New List(Of BlockReference)
            Dim EntitysPlanta As New List(Of Entity)
            Dim XRefPlanta As BlockReference = Nothing
            Dim XRefMode As BlockReference = Nothing
            Dim SelectedLayout As String
            Dim Position As Point3d
            Dim TaskList As List(Of String)
            Dim LayoutList As List(Of Layout)
            Dim frmTaskList As frmTaskList = Nothing
            Dim LayoutManager As LayoutManager = LayoutManager.Current
            Dim DefaultTask As String
            Dim Response As Object
            Dim ProgressMeter As ProgressMeter = Nothing
            Dim ViewTableRecord As ViewTableRecord
            Dim Center3d As Point3d
            Dim Center As Point2d
            Dim Height As Double
            Dim Width As Double
            Dim Extents3d As Extents3d
            Dim Entity As Entity
            Dim DBObjectCollection As New DBObjectCollection
            Dim Matrix3d As Matrix3d
            Dim BlockTableRecord As BlockTableRecord
            Dim XrefAttNames As List(Of String)
            Dim AlimMode As Object
            Dim AlimPlan As Object
            Dim FoundBlocks As New List(Of BlockReference)
            Dim ScaleFactor As Double
            Dim LineAmount As Integer
            Dim BlockName As String
            Dim ListofBlockNames As List(Of String)
            Dim DR() As DataRow

            Try

                'Obtem os dados de configuração
                Using DB As New adoNetExtension.AdoNetConnect.AdoNet
                    With DB
                        .OpenConnection(MyPlugin.StringConnection, adoNetExtension.AdoNetConnect.AdoNet.Providers.SQL)
                        With .SqlSelect
                            .Clear()
                            .AppendLine("SELECT")
                            .AppendLine("CON.conID")
                            .AppendLine(", CON.conBlockName")
                            .AppendLine(", CON.conScaleXP")
                            .AppendLine(", CON.conLines")
                            .AppendLine(", CON.ATIVO")
                            .AppendLine("FROM")
                            .AppendLine("ragaConfig AS CON")
                            .AppendLine("WHERE")
                            .AppendLine("(")
                            .AppendLine("CON.ATIVO = 1")
                            .AppendLine(")")
                            _DtConfig = .CreateConsult
                        End With
                        .CloseConnection()
                    End With
                End Using

                'Avalia
                If _DtConfig.Rows.Count = 0 Then
                    MsgBox("Nenhuma configuração foi especificada para o comando.", MsgBoxStyle.Exclamation)
                    Exit Sub
                End If

                ListofBlockNames = New List(Of String)

                For i As Integer = 0 To _DtConfig.Rows.Count - 1

                    ListofBlockNames.Add(_DtConfig.Rows(i).Item("conBlockName"))

                Next

                ' DR = _DtConfig.Rows(0)



                'Moda para 'Model' se estiver no 'Paper'
                If LayoutManager.CurrentLayout <> "Model" Then
                    LayoutManager.CurrentLayout = "Model"
                End If

                'Solicita a seleção de planta
                Entity = Engine2.AcadInterface.GetEntity("Selecione o XREF com a planta baixa e alimentadores",,,,,, New List(Of Type)({GetType(BlockReference)}),)

                'Avalia
                If IsNothing(Entity) = False Then

                    'Obtem o bloco
                    XRefPlanta = Entity

                    'Instancia a classe de seleção em blocos
                    _SelectEntityInBlockPlane = New SelectEntityInBlockPlane

                    'Solicita a seleção no XRef
                    EntitysPlanta = _SelectEntityInBlockPlane.GetSelection(XRefPlanta)

                    'Avalia
                    If EntitysPlanta.Count > 0 Then

                        'Avalia se é XREF
                        If Engine2.Block.IsXref(XRefPlanta.Name) = True Then

                            'Solicita a seleção de alimentadores
                            Entity = Engine2.AcadInterface.GetEntity("Selecione o XREF com alimentadores do MODE",,,,,, New List(Of Type)({GetType(BlockReference)}),)

                            'Avalia
                            If IsNothing(Entity) = False Then

                                'Obtem o bloco
                                XRefMode = Entity

                                'Avalia se é XREF
                                If Engine2.Block.IsXref(XRefMode.Name) = True Then

                                    'Carrega os nomes dos espaços de desenho
                                    LayoutList = Engine2.AcadLayout.ListLayouts(Document)
                                    TaskList = New List(Of String)
                                    For Each Layout As Layout In LayoutList
                                        If Layout.LayoutName <> "Model" Then
                                            TaskList.Add(Layout.LayoutName)
                                        End If
                                    Next

                                    'Obtem o espaço corrente
                                    DefaultTask = TaskList(0)

                                    'Avalia a quantidade
                                    If TaskList.Count > 1 Then

                                        'Instancia a janela de seleção
                                        frmTaskList = New frmTaskList

                                        'Abre a janela para seleção
                                        SelectedLayout = frmTaskList.ShowDialog("Selecione o Layout", TaskList, DefaultTask)

                                    Else

                                        'Seleciona o layout único
                                        SelectedLayout = DefaultTask

                                    End If

                                    'Avalia
                                    Select Case SelectedLayout
                                        Case Nothing

                                            'Mensagem
                                            Editor.WriteMessage(vbLf & "Fim do comando" & vbLf)
                                            'Finaliza
                                            Exit Sub
                                        Case Else

                                            'Selecionar o layout escolhido
                                            If SelectedLayout <> LayoutManager.CurrentLayout Then
                                                LayoutManager.CurrentLayout = SelectedLayout
                                            End If

                                    End Select

                                    'Solicita o ponto de inserção
                                    Response = Engine2.AcadInterface.GetPoint("Selecione o ponto de inserção da lista:")

                                    'Avalia
                                    If IsNothing(Response) = False Then

                                        'Obtém o ponto
                                        Position = Response

                                        'Bloquei edição
                                        Using Editor.Document.LockDocument

                                            'Transação
                                            Using Transaction As Transaction = Database.TransactionManager.StartTransaction()

                                                Try

                                                    'Abre o temporizador
                                                    ProgressMeter = New ProgressMeter
                                                    ProgressMeter.SetLimit(3)
                                                    ProgressMeter.Start("Analisando o projeto, aguarde...")

                                                    'Atualiza os XRefs
                                                    Engine2.Block.ReloadXrefs(New List(Of String)({XRefPlanta.Name, XRefMode.Name}), Transaction)

                                                    'Atualiza o temporizador
                                                    ProgressMeter.MeterProgress()

                                                    'Obtém os dados do Mode
                                                    '----------------------------------------------------

                                                    'Obtém BlockTableRecord
                                                    BlockTableRecord = Transaction.GetObject(XRefMode.BlockTableRecord, OpenMode.ForRead)

                                                    'Obtém os blocos válidos (Alimentadores)
                                                    BlocksMode = New List(Of BlockReference)
                                                    For Each id As ObjectId In BlockTableRecord
                                                        Entity = Transaction.GetObject(id, OpenMode.ForRead)
                                                        If Entity.IsErased = False Then
                                                            If Entity.GetType = GetType(BlockReference) AndAlso
                                                               CType(Entity, BlockReference).Name.Contains("SPHE-EFO-ALIMENTADOR_") = True Then
                                                                BlocksMode.Add(Entity)
                                                            End If
                                                        End If
                                                    Next

                                                    'Atualiza o temporizador
                                                    ProgressMeter.MeterProgress()

                                                    'Obtém os dados da planta
                                                    '----------------------------------------------------

                                                    'Obtém BlockTableRecord
                                                    BlockTableRecord = Transaction.GetObject(XRefPlanta.BlockTableRecord, OpenMode.ForRead)

                                                    'Obtém os blocos válidos (XRef)
                                                    BlocksPlanta = New List(Of BlockReference)
                                                    For Each EntitySel As Entity In EntitysPlanta
                                                        For Each id As ObjectId In BlockTableRecord
                                                            Entity = Transaction.GetObject(id, OpenMode.ForRead)
                                                            If EntitySel.Handle.ToString = Entity.Handle.ToString Then
                                                                BlocksPlanta.Add(Entity)
                                                            End If
                                                        Next
                                                    Next

                                                    'Atualiza o temporizador
                                                    ProgressMeter.MeterProgress()
                                                    ProgressMeter.Stop()
                                                    ProgressMeter.Dispose()

                                                    'Gera a lista
                                                    '----------------------------------------------------

                                                    'Abre o temporizador
                                                    ProgressMeter = New ProgressMeter
                                                    ProgressMeter.SetLimit(BlocksMode.Count + BlocksPlanta.Count)
                                                    ProgressMeter.Start("Comparando alimentadores, aguarde...")

                                                    'Cria a coleção de nomes dos alimentadores processados
                                                    XrefAttNames = New List(Of String)

                                                    'Obtém a lista de blocos válidos
                                                    For Each BLKMODE As BlockReference In BlocksMode
                                                        AlimMode = Engine2.Block.ReadAttribute(BLKMODE, "1TP1", Transaction)
                                                        If IsNothing(AlimMode) = False Then
                                                            For Each BLKPLAN As BlockReference In BlocksPlanta
                                                                DR = _DtConfig.Select("conBlockName = '" & BLKPLAN.Name & "'")

                                                                Dim ScaleXP As Double = DR(0).Item("conScaleXP")

                                                                AlimPlan = Engine2.Block.ReadAttribute(BLKPLAN, "XXXX", Transaction)
                                                                If IsNothing(AlimPlan) = False Then
                                                                    If AlimMode.ToString.Equals(AlimPlan.ToString, StringComparison.OrdinalIgnoreCase) = True Then
                                                                        If XrefAttNames.Contains(AlimMode) = False Then
                                                                            FoundBlocks.Add(BLKMODE)
                                                                            XrefAttNames.Add(AlimMode)
                                                                        End If
                                                                    End If
                                                                End If
                                                                'Atualiza o temporizador
                                                                ProgressMeter.MeterProgress()
                                                            Next
                                                            'Atualiza o temporizador
                                                            ProgressMeter.MeterProgress()
                                                        End If
                                                    Next

                                                    'Atualiza o temporizador
                                                    ProgressMeter.Stop()
                                                    ProgressMeter.Dispose()

                                                    'Ordena os blocos a partir de sua posição em Y
                                                    FoundBlocks.Sort(Function(X As BlockReference, Y As BlockReference) Y.Position.Y.CompareTo(X.Position.Y))

                                                    'Obtém a matrix de deslocamento a partir da coordenada 0,0,0 do XRef
                                                    Matrix3d = Matrix3d.Displacement(New Point3d(0, 0, 0).GetVectorTo(XRefMode.Position))

                                                    'Abre o temporizador
                                                    ProgressMeter = New ProgressMeter
                                                    ProgressMeter.SetLimit(FoundBlocks.Count)
                                                    ProgressMeter.Start("Criando viewports, aguarde...")

                                                    'Percorre a coleção
                                                    For Each BlockReference As BlockReference In FoundBlocks

                                                        'Obtém os dados dimensionais
                                                        Extents3d = BlockReference.GeometricExtents

                                                        'Obtém o comprimento do bloco
                                                        Width = New Point2d(Extents3d.MinPoint.X, 0).GetDistanceTo(New Point2d(Extents3d.MaxPoint.X, 0))

                                                        'Obtém a altura do bloco
                                                        Height = New Point2d(0, Extents3d.MaxPoint.Y).GetDistanceTo(New Point2d(0, Extents3d.MinPoint.Y))

                                                        'Cria o ponto 3d com base no 2d
                                                        Center3d = Engine2.Geometry.MidPoint(Extents3d.MinPoint, Extents3d.MaxPoint)

                                                        'Constrói a visualização
                                                        ViewTableRecord = New ViewTableRecord

                                                        'Recalcula o ponto com base na matriz
                                                        Center3d = Center3d.TransformBy(Matrix3d)

                                                        'Obtém o ponto 2d
                                                        Center = Engine2.Geometry.Point3dToPoint2d(Center3d)

                                                        'Define o ponto central
                                                        ViewTableRecord.CenterPoint = Center

                                                        'Define a altura
                                                        ViewTableRecord.Height = Height

                                                        'Define o comprimento
                                                        ViewTableRecord.Width = Width

                                                        'Seta a direção
                                                        ViewTableRecord.ViewDirection = New Vector3d(0, 0, 1)

                                                        'Aplica escala no comprimento
                                                        Width = Width * ScaleFactor

                                                        'Aplica escala na altura
                                                        Height = Height * ScaleFactor

                                                        'Cria a viewport
                                                        Engine2.AcadViewport.CreateViewport(LayoutManager.CurrentLayout, Position, Width, Height, ViewTableRecord, ScaleFactor, True, True, False,, Transaction)

                                                        Position = New Point3d(Position.X, Position.Y - Height, Position.Z)

                                                        'Atualiza o temporizador
                                                        ProgressMeter.MeterProgress()

                                                    Next

                                                    'Finaliza o temporizador                                                                                        
                                                    ProgressMeter.Stop()
                                                    'Confirma a transação
                                                    Transaction.Commit()

                                                    'Mensagem
                                                    Editor.WriteMessage(vbLf & "Fim do comando" & vbLf)

                                                Catch ex As System.Exception

                                                    'Aborta a transação
                                                    Transaction.Abort()
                                                    Throw New System.Exception(ex.Message)

                                                End Try

                                            End Using

                                        End Using

                                    End If

                                End If

                            End If

                        End If

                    End If

                End If

                'Mensagem
                Editor.WriteMessage(vbLf & "Fim do comando" & vbLf)

            Catch ex As System.Exception

                'Erro
                MsgBox(ex.Message, MsgBoxStyle.Critical)

            Finally

                'Destrói a janela 
                If IsNothing(frmTaskList) = False Then
                    frmTaskList.Dispose()
                End If

                'Destrói o temporizador
                If IsNothing(ProgressMeter) = False Then
                    ProgressMeter.Dispose()
                End If

            End Try

        End Sub

    End Class

End Namespace
