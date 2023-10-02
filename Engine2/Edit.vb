Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports System.Text
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports ElectricTools2020.Engine2.AcadInterface

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    Public Class Edit

        ''' <summary>
        ''' Junta os segmentos (Line, Curve, Arc e Polyline) associados a uma polyline
        ''' </summary>
        ''' <param name="Transaction">Transaction</param>
        ''' <param name="Polyline">Polyline inicial</param>
        ''' <param name="FilterSelectSequenceOptions">Opções de filtro para seleção</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function JoinSelection(Transaction As Transaction, Polyline As Polyline, FilterSelectSequenceOptions As FilterSelectSequenceOptions) As Polyline
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim SelectSequenceResult As SelectSequenceResult
            Dim DBObjectCollection As New DBObjectCollection
            Try
                SelectSequenceResult = Engine2.AcadInterface.SelectCurveSequence(Polyline, FilterSelectSequenceOptions)
                If SelectSequenceResult.Count <> 0 Then
                    For Each SelectSequenceItem As SelectSequenceItem In SelectSequenceResult
                        If DBObjectCollection.Contains(SelectSequenceItem.BaseEntity) = False Then
                            DBObjectCollection.Add(SelectSequenceItem.BaseEntity)
                        End If
                        If DBObjectCollection.Contains(SelectSequenceItem.AssocEntity) = False Then
                            DBObjectCollection.Add(SelectSequenceItem.AssocEntity)
                        End If
                    Next
                    Polyline = Engine2.EntityInteration.Join(Transaction, DBObjectCollection)
                End If
                Return Polyline
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Quebra a entidade no ponto especificado e retorna a entidade resultante
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Point3dCollection">Coleção de pontos para quebra</param>
        ''' <param name="Message">Mensagem de conclusão (Prompt)</param>
        ''' <returns>Nothing\ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function Break(Entity As Entity, Point3dCollection As Point3dCollection, Optional Message As Boolean = True) As Object
            Try
                Dim ObjectIdCollection As New ObjectIdCollection
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
                Dim BlockTableRecord As BlockTableRecord
                Dim DBObjectCollection As DBObjectCollection
                Dim lstDouble As List(Of Double)
                Dim Ent As Object = Nothing
                If Point3dCollection.Count = 0 Then
                    Editor.WriteMessage(vbLf & "Nenhum ponto de quebra foi informado" & vbLf)
                    Return Nothing
                Else
                    If Engine2.EntityInteration.PointAtEntity(Entity, Point3dCollection) = False Then
                        Editor.WriteMessage(vbLf & "O ponto informado não pertence a entidade" & vbLf)
                        Return Nothing
                    Else
                        Select Case Entity.GetType.Name
                            Case "Line"
                                Ent = New Line
                            Case "Spline"
                                Ent = New Spline
                            Case "Arc"
                                Ent = New Arc
                            Case "Polyline"
                                Ent = New Polyline
                            Case "Polyline2d"
                                Ent = New Polyline2d
                            Case "Polyline3d"
                                Ent = New Polyline3d
                            Case Else
                                If Message = True Then
                                    Editor.WriteMessage(vbLf & "Entidade do tipo " & Entity.GetType.Name & " não é válida" & vbLf)
                                End If
                                Return Nothing
                        End Select
                        Using Editor.Document.LockDocument
                            Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                                Ent = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite)
                                lstDouble = New List(Of Double)
                                For Each Point3d As Point3d In Point3dCollection
                                    lstDouble.Add(Ent.GetParameterAtPoint(Ent.GetClosestPointTo(Point3d, False)))
                                Next
                                lstDouble.Sort()
                                DBObjectCollection = Ent.GetSplitCurves(New DoubleCollection(lstDouble.ToArray()))
                                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                                Try
                                    For Each EntityItem As Entity In DBObjectCollection
                                        BlockTableRecord.AppendEntity(EntityItem)
                                        EntityItem.SetDatabaseDefaults()
                                        Transaction.AddNewlyCreatedDBObject(EntityItem, True)
                                        ObjectIdCollection.Add(EntityItem.ObjectId)
                                    Next
                                    Ent.UpgradeOpen()
                                    Ent.Erase()
                                    Ent.Dispose()
                                    Transaction.Commit()
                                Catch
                                    Transaction.Abort()
                                End Try
                            End Using
                        End Using
                        If ObjectIdCollection.Count > 1 Then
                            Return ObjectIdCollection
                        Else
                            Return Nothing
                        End If
                    End If
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Quebra a entidade no ponto especificado e retorna a entidade resultante (Para uso em lote)
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Point3dCollection">Coleção de pontos para quebra</param>
        ''' <param name="Message">Mensagem de conclusão (Prompt)</param>
        ''' <param name="EraseOriginalEntity">Determina se o segmento original deve ser excluído</param>
        ''' <returns>Nothing\ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function Break(Transaction As Transaction, Entity As Entity, Point3dCollection As Point3dCollection, Optional Message As Boolean = True, Optional EraseOriginalEntity As Boolean = False) As Object
            Try
                Dim ObjectIdCollection As New ObjectIdCollection
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
                Dim BlockTableRecord As BlockTableRecord
                Dim DBObjectCollection As DBObjectCollection
                Dim lstDouble As List(Of Double)
                Dim Ent As Object = Nothing
                If Point3dCollection.Count = 0 Then
                    Editor.WriteMessage(vbLf & "Nenhum ponto de quebra foi informado" & vbLf)
                    Return Nothing
                Else
                    If Engine2.EntityInteration.PointAtEntity(Entity, Point3dCollection) = False Then
                        Editor.WriteMessage(vbLf & "O ponto informado não pertence a entidade" & vbLf)
                        Return Nothing
                    Else
                        Select Case Entity.GetType.Name
                            Case "Line"
                                Ent = New Line
                            Case "Spline"
                                Ent = New Spline
                            Case "Arc"
                                Ent = New Arc
                            Case "Polyline"
                                Ent = New Polyline
                            Case "Polyline2d"
                                Ent = New Polyline2d
                            Case "Polyline3d"
                                Ent = New Polyline3d
                            Case Else
                                If Message = True Then
                                    Editor.WriteMessage(vbLf & "Entidade do tipo " & Entity.GetType.Name & " não é válida" & vbLf)
                                End If
                                Return Nothing
                        End Select
                        Ent = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite)
                        lstDouble = New List(Of Double)
                        For Each Point3d As Point3d In Point3dCollection
                            lstDouble.Add(Ent.GetParameterAtPoint(Ent.GetClosestPointTo(Point3d, False)))
                        Next
                        lstDouble.Sort()
                        DBObjectCollection = Ent.GetSplitCurves(New DoubleCollection(lstDouble.ToArray()))
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        Try
                            For Each EntityItem As Entity In DBObjectCollection
                                BlockTableRecord.AppendEntity(EntityItem)
                                EntityItem.SetDatabaseDefaults()
                                Transaction.AddNewlyCreatedDBObject(EntityItem, True)
                                ObjectIdCollection.Add(EntityItem.ObjectId)
                            Next
                            If EraseOriginalEntity = True Then
                                Ent.UpgradeOpen()
                                Ent.Erase()
                                Ent.Dispose()
                            End If
                        Catch ex As System.Exception
                            Throw New System.Exception("Erro em Edit.Break, motivo: " & ex.Message)
                        End Try
                        If ObjectIdCollection.Count > 1 Then
                            Return ObjectIdCollection
                        Else
                            Return Nothing
                        End If
                    End If
                End If
            Catch
                Return Nothing
            End Try
        End Function


        ''' <summary>
        ''' Quebra a entidade no ponto especificado e retorna a entidade resultante
        ''' </summary>
        ''' <param name="ObjectId">ID da entidade</param>
        ''' <param name="Point3dCollection">Coleção de pontos para quebra</param>
        ''' <param name="Message">Mensagem de conclusão (Prompt)</param>
        ''' <returns>Nothing\ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function Break(ObjectId As ObjectId, Point3dCollection As Point3dCollection, Optional Message As Boolean = True) As Object
            Return Engine2.Edit.Break(Engine2.ConvertObject.ObjectIDToEntity(ObjectId), Point3dCollection, Message)
        End Function

        ''' <summary>
        ''' Explode uma entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Explode(Entity As Entity) As Boolean
            Try
                Dim Database As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
                Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
                Dim DBObjectCollection As New DBObjectCollection
                Dim BlockTableRecord As BlockTableRecord
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        BlockTableRecord = DirectCast(Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                        Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
                        Entity.UpgradeOpen()
                        Entity.Explode(DBObjectCollection)
                        For Each DBObject As DBObject In DBObjectCollection
                            BlockTableRecord.AppendEntity(DBObject)
                            Transaction.AddNewlyCreatedDBObject(DBObject, True)
                        Next
                        Entity.UpgradeOpen()
                        Entity.Erase()
                        Entity.Dispose()
                        Transaction.Commit()
                    End Using
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Explode a entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="EraseOriginalEntity">Determina se a entidade original deve ser excluída</param>
        ''' <param name="NotExplosivedToClone">Determina se entidades que não possam ser explodidas serão incluídas em cópia na coleção do usuário</param>
        ''' <param name="SubItems">Determina se os subitens serão retornados em caso de entidades complexas</param>
        ''' <param name="NestedEntitys">Determina se serão retornados apenas entidades que não podem mais ser explodidas em todos os subitens</param>
        ''' <returns>Nothing\DBObjectCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function Explode(Entity As Entity, Optional EraseOriginalEntity As Boolean = False, Optional NotExplosivedToClone As Boolean = False, Optional SubItems As Boolean = False, Optional NestedEntitys As Boolean = False) As Object
            Dim Database As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim TempDBObjectCollection As New DBObjectCollection
            Dim DBObjectCollection As New DBObjectCollection
            Dim FilterDBObjectCollection As New DBObjectCollection
            Dim BlockTableRecord As BlockTableRecord
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                    Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
                    Entity.UpgradeOpen()
                    Try
                        Entity.Explode(TempDBObjectCollection)
                    Catch
                        If NotExplosivedToClone = True Then
                            TempDBObjectCollection.Add(Entity.Clone())
                        End If
                    End Try
                    If EraseOriginalEntity = True Then
                        Entity.UpgradeOpen()
                        Entity.Erase()
                        Entity.Dispose()
                    End If
                    If SubItems = True Then
                        While TempDBObjectCollection.Count <> 0
                            Entity = TempDBObjectCollection.Item(0)
                            Try
                                FilterDBObjectCollection.Clear()
                                Entity.Explode(FilterDBObjectCollection)
                                For Each DBObject As DBObject In FilterDBObjectCollection
                                    Entity = DBObject
                                    If NestedEntitys = True Then
                                        If Engine2.EntityInteration.IsExplosible(Entity) = True Then
                                            TempDBObjectCollection.Add(Entity)
                                        Else
                                            DBObjectCollection.Add(Entity)
                                        End If
                                    Else
                                        DBObjectCollection.Add(Entity)
                                    End If
                                Next
                            Catch
                                DBObjectCollection.Add(Entity)
                            End Try
                            TempDBObjectCollection.RemoveAt(0)
                        End While
                    Else
                        For Each DBObject As DBObject In TempDBObjectCollection
                            DBObjectCollection.Add(DBObject)
                        Next
                    End If
                    For Each DBObject As DBObject In DBObjectCollection
                        BlockTableRecord.AppendEntity(DBObject)
                        Transaction.AddNewlyCreatedDBObject(DBObject, True)
                    Next
                    Transaction.Commit()
                End Using
            End Using
            If DBObjectCollection.Count = 0 Then
                Return Nothing
            Else
                Return DBObjectCollection
            End If
        End Function

        ''' <summary>
        ''' Explode a entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="DBObjectCollection">Item para armazenagem das partes da explosão (Acumulativo)</param>
        ''' <param name="EraseOriginalEntity">Determina se a entidade original deve ser excluída</param>
        ''' <param name="NotExplosivedToClone">Determina se entidades que não possam ser explodidas serão incluídas em cópia na coleção do usuário</param>
        ''' <param name="SubItems">Determina se os subitens serão retornados em caso de entidades complexas</param>
        ''' <param name="NestedEntitys">Determina se serão retornados apenas entidades que não podem mais ser explodidas em todos os subitens</param>
        ''' <remarks></remarks>
        Public Shared Sub Explode(Entity As Entity, ByRef DBObjectCollection As DBObjectCollection, Optional EraseOriginalEntity As Boolean = False, Optional NotExplosivedToClone As Boolean = False, Optional SubItems As Boolean = False, Optional NestedEntitys As Boolean = False)
            Dim Database As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim TempDBObjectCollection As New DBObjectCollection
            Dim SubItemsDBObjectCollection As New DBObjectCollection
            Dim FilterDBObjectCollection As New DBObjectCollection
            Dim BlockTableRecord As BlockTableRecord
            If IsNothing(DBObjectCollection) = True Then
                DBObjectCollection = New DBObjectCollection
            End If
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                    Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
                    Entity.UpgradeOpen()
                    Try
                        Entity.Explode(TempDBObjectCollection)
                    Catch
                        If NotExplosivedToClone = True Then
                            TempDBObjectCollection.Add(Entity.Clone())
                        End If
                    End Try
                    If EraseOriginalEntity = True Then
                        Entity.UpgradeOpen()
                        Entity.Erase()
                        Entity.Dispose()
                    End If
                    If SubItems = True Then
                        While TempDBObjectCollection.Count <> 0
                            Entity = TempDBObjectCollection.Item(0)
                            Try
                                FilterDBObjectCollection.Clear()
                                Entity.Explode(FilterDBObjectCollection)
                                For Each DBObject As DBObject In FilterDBObjectCollection
                                    Entity = DBObject
                                    If NestedEntitys = True Then
                                        If Engine2.EntityInteration.IsExplosible(Entity) = True Then
                                            TempDBObjectCollection.Add(Entity)
                                        Else
                                            SubItemsDBObjectCollection.Add(Entity)
                                        End If
                                    Else
                                        SubItemsDBObjectCollection.Add(Entity)
                                    End If
                                Next
                            Catch
                                SubItemsDBObjectCollection.Add(Entity)
                            End Try
                            TempDBObjectCollection.RemoveAt(0)
                        End While
                    Else
                        SubItemsDBObjectCollection = TempDBObjectCollection
                    End If
                    For Each DBObject As DBObject In SubItemsDBObjectCollection
                        BlockTableRecord.AppendEntity(DBObject)
                        Transaction.AddNewlyCreatedDBObject(DBObject, True)
                        DBObjectCollection.Add(DBObject)
                    Next
                    Transaction.Commit()
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Explode a entidade
        ''' </summary>
        ''' <param name="ObjectIdCollection">Coleção de ID´s a serem explodidos</param>
        ''' <param name="EraseOriginalEntity">Determina se a entidade original deve ser excluída</param>
        ''' <param name="NotExplosivedToClone">Determina se entidades que não possam ser explodidas serão incluídas em cópia na coleção do usuário</param>
        ''' <param name="SubItems">Determina se os subitens serão retornados em caso de entidades complexas</param>
        ''' <param name="NestedEntitys">Determina se serão retornados apenas entidades que não podem mais ser explodidas em todos os subitens</param>
        ''' <returns>Nothing\DBObjectCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function Explode(ObjectIdCollection As ObjectIdCollection, Optional EraseOriginalEntity As Boolean = False, Optional NotExplosivedToClone As Boolean = False, Optional SubItems As Boolean = False, Optional NestedEntitys As Boolean = False) As Object
            Dim Database As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Entity As Entity
            Dim TempDBObjectCollection As New DBObjectCollection
            Dim DBObjectCollection As New DBObjectCollection
            Dim FilterDBObjectCollection As New DBObjectCollection
            Dim BlockTableRecord As BlockTableRecord
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    For Each ObjectId As ObjectId In ObjectIdCollection
                        Entity = Transaction.GetObject(ObjectId, OpenMode.ForRead)
                        Entity.UpgradeOpen()
                        Try
                            Entity.Explode(TempDBObjectCollection)
                        Catch
                            If NotExplosivedToClone = True Then
                                TempDBObjectCollection.Add(Entity.Clone())
                            End If
                        End Try
                        If EraseOriginalEntity = True Then
                            Entity.UpgradeOpen()
                            Entity.Erase()
                            Entity.Dispose()
                        End If
                    Next
                    If SubItems = True Then
                        While TempDBObjectCollection.Count <> 0
                            Entity = TempDBObjectCollection.Item(0)
                            Try
                                FilterDBObjectCollection.Clear()
                                Entity.Explode(FilterDBObjectCollection)
                                For Each DBObject As DBObject In FilterDBObjectCollection
                                    Entity = DBObject
                                    If NestedEntitys = True Then
                                        If Engine2.EntityInteration.IsExplosible(Entity) = True Then
                                            TempDBObjectCollection.Add(Entity)
                                        Else
                                            DBObjectCollection.Add(Entity)
                                        End If
                                    Else
                                        DBObjectCollection.Add(Entity)
                                    End If
                                Next
                            Catch
                                DBObjectCollection.Add(Entity)
                            End Try
                            TempDBObjectCollection.RemoveAt(0)
                        End While
                    Else
                        DBObjectCollection = TempDBObjectCollection
                    End If
                    BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                    For Each DBObject As DBObject In DBObjectCollection
                        BlockTableRecord.AppendEntity(DBObject)
                        Transaction.AddNewlyCreatedDBObject(DBObject, True)
                        DBObjectCollection.Add(DBObject)
                    Next
                    Transaction.Commit()
                End Using
            End Using
            If DBObjectCollection.Count = 0 Then
                Return Nothing
            Else
                Return DBObjectCollection
            End If
        End Function

        ''' <summary>
        ''' Explode a entidade
        ''' </summary>
        ''' <param name="ObjectIdCollection">Coleção de ID´s a serem explodidos</param>
        ''' <param name="DBObjectCollection">Item para armazenagem das partes da explosão (Acumulativo)</param>
        ''' <param name="EraseOriginalEntity">Determina se a entidade original deve ser excluída</param>
        ''' <param name="NotExplosivedToClone">Determina se entidades que não possam ser explodidas serão incluídas em cópia na coleção do usuário</param>
        ''' <param name="SubItems">Determina se os subitens serão retornados em caso de entidades complexas</param>
        ''' <param name="NestedEntitys">Determina se serão retornados apenas entidades que não podem mais ser explodidas em todos os subitens</param>
        ''' <param name="EntitysException">O tipo de entidade especificada será desconsiderada para SubItems e NestedEntitys</param>
        ''' <remarks></remarks>
        Public Shared Sub Explode(ObjectIdCollection As ObjectIdCollection, ByRef DBObjectCollection As DBObjectCollection, Optional EraseOriginalEntity As Boolean = False, Optional NotExplosivedToClone As Boolean = False, Optional SubItems As Boolean = False, Optional NestedEntitys As Boolean = False, Optional EntitysException As ArrayList = Nothing)
            Dim Database As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Entity As Entity
            Dim TempDBObjectCollection As New DBObjectCollection
            Dim SubItemsDBObjectCollection As New DBObjectCollection
            Dim FilterDBObjectCollection As New DBObjectCollection
            Dim BlockTableRecord As BlockTableRecord
            Dim ProgressMeter As ProgressMeter = Nothing
            Try
                If IsNothing(DBObjectCollection) = True Then
                    DBObjectCollection = New DBObjectCollection
                End If
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        ProgressMeter = New ProgressMeter
                        ProgressMeter.SetLimit(ObjectIdCollection.Count)
                        ProgressMeter.Start("Analisando o arquivo, aguarde...")
                        System.Windows.Forms.Application.DoEvents()
                        For Each ObjectId As ObjectId In ObjectIdCollection
                            Entity = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                            Entity.UpgradeOpen()
                            Try
                                Entity.Explode(TempDBObjectCollection)
                            Catch
                                If NotExplosivedToClone = True Then
                                    TempDBObjectCollection.Add(Entity.Clone())
                                End If
                            End Try
                            If EraseOriginalEntity = True Then
                                Entity.UpgradeOpen()
                                Entity.Erase()
                                Entity.Dispose()
                            End If
                            ProgressMeter.MeterProgress()
                            System.Windows.Forms.Application.DoEvents()
                        Next
                        ProgressMeter.Stop()
                        System.Windows.Forms.Application.DoEvents()
                        If SubItems = True Then
                            ProgressMeter.SetLimit(TempDBObjectCollection.Count)
                            ProgressMeter.Start("Analisando sub entidades, aguarde...")
                            System.Windows.Forms.Application.DoEvents()
                            While TempDBObjectCollection.Count <> 0
                                Entity = TempDBObjectCollection.Item(0)
                                If IsNothing(EntitysException) = False Then
                                    If EntitysException.Contains(Entity.GetType.Name) = False Then
                                        Try
                                            FilterDBObjectCollection.Clear()
                                            Entity.Explode(FilterDBObjectCollection)
                                            For Each DBObject As DBObject In FilterDBObjectCollection
                                                Entity = DBObject
                                                If NestedEntitys = True Then
                                                    If Engine2.EntityInteration.IsExplosible(Entity, EntitysException) = True Then
                                                        TempDBObjectCollection.Add(Entity)
                                                    Else
                                                        SubItemsDBObjectCollection.Add(Entity)
                                                    End If
                                                Else
                                                    SubItemsDBObjectCollection.Add(Entity)
                                                End If
                                            Next
                                        Catch
                                            SubItemsDBObjectCollection.Add(Entity)
                                        End Try
                                    End If
                                    TempDBObjectCollection.RemoveAt(0)
                                    ProgressMeter.MeterProgress()
                                    System.Windows.Forms.Application.DoEvents()
                                Else
                                    Try
                                        FilterDBObjectCollection.Clear()
                                        Entity.Explode(FilterDBObjectCollection)
                                        For Each DBObject As DBObject In FilterDBObjectCollection
                                            Entity = DBObject
                                            If NestedEntitys = True Then
                                                If Engine2.EntityInteration.IsExplosible(Entity) = True Then
                                                    TempDBObjectCollection.Add(Entity)
                                                Else
                                                    SubItemsDBObjectCollection.Add(Entity)
                                                End If
                                            Else
                                                SubItemsDBObjectCollection.Add(Entity)
                                            End If
                                        Next
                                    Catch
                                        SubItemsDBObjectCollection.Add(Entity)
                                    End Try
                                    TempDBObjectCollection.RemoveAt(0)
                                    ProgressMeter.MeterProgress()
                                    System.Windows.Forms.Application.DoEvents()
                                End If
                            End While
                            ProgressMeter.Stop()
                            System.Windows.Forms.Application.DoEvents()
                        Else
                            SubItemsDBObjectCollection = TempDBObjectCollection
                        End If
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        ProgressMeter.SetLimit(SubItemsDBObjectCollection.Count)
                        ProgressMeter.Start("Explodindo entidades, aguarde...")
                        System.Windows.Forms.Application.DoEvents()
                        For Each DBObject As DBObject In SubItemsDBObjectCollection
                            Entity = DBObject
                            BlockTableRecord.AppendEntity(Entity)
                            Transaction.AddNewlyCreatedDBObject(Entity, True)
                            DBObjectCollection.Add(Entity)
                            ProgressMeter.MeterProgress()
                            System.Windows.Forms.Application.DoEvents()
                        Next
                        ProgressMeter.Stop()
                        System.Windows.Forms.Application.DoEvents()
                        Transaction.Commit()
                    End Using
                End Using
            Finally
                If IsNothing(ProgressMeter) = False Then
                    ProgressMeter.Stop()
                    ProgressMeter.Dispose()
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Exclui a entidade
        ''' </summary>
        ''' <param name="Entity">Entity</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EraseEntity(ByVal Entity As Entity) As Boolean
            Try
                Dim Database As Database = HostApplicationServices.WorkingDatabase
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Using Editor.Document.LockDocument
                    Using Transaction = Database.TransactionManager.StartTransaction()
                        Entity = DirectCast(Transaction.GetObject(Entity.Id, OpenMode.ForWrite), Entity)
                        'Entity.UpgradeOpen()
                        Entity.Erase()
                        Entity.Dispose()
                        Transaction.Commit()
                    End Using
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Exclui a entidade
        ''' </summary>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EraseEntity(ByVal ObjectId As ObjectId) As Boolean
            Try
                Dim Database As Database = HostApplicationServices.WorkingDatabase
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Using Editor.Document.LockDocument
                    Using Transaction = Database.TransactionManager.StartTransaction()
                        Dim BlockTable As BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                        Dim Entity As Entity = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                        'Entity.UpgradeOpen()
                        Entity.Erase()
                        Entity.Dispose()
                        Transaction.Commit()
                    End Using
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Exclui a entidade
        ''' </summary>
        ''' <param name="DBObject">DBObject</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EraseEntity(ByVal DBObject As DBObject) As Boolean
            Try
                Dim Database As Database = HostApplicationServices.WorkingDatabase
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Using Editor.Document.LockDocument
                    Using Transaction = Database.TransactionManager.StartTransaction()
                        Dim Entity As Entity = DirectCast(Transaction.GetObject(DBObject.Id, OpenMode.ForWrite), Entity)
                        'Entity.UpgradeOpen()
                        Entity.Erase()
                        Entity.Dispose()
                        Transaction.Commit()
                    End Using
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Exclui a entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Entity">Entity</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EraseEntity(Transaction As Transaction, ByVal Entity As Entity) As Boolean
            Try
                Dim Database As Database = HostApplicationServices.WorkingDatabase
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Entity = Transaction.GetObject(Entity.Id, OpenMode.ForWrite)
                'Entity.UpgradeOpen()
                Entity.Erase()
                Entity.Dispose()
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Exclui a entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EraseEntity(Transaction As Transaction, ByVal ObjectId As ObjectId) As Boolean
            Try
                Dim Database As Database = HostApplicationServices.WorkingDatabase
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Dim Entity As Entity = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                'Entity.UpgradeOpen()
                Entity.Erase()
                Entity.Dispose()
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Exclui a entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DBObject">DBObject</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EraseEntity(Transaction As Transaction, ByVal DBObject As DBObject) As Boolean
            Try
                Dim Database As Database = HostApplicationServices.WorkingDatabase
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Dim Entity As Entity = DirectCast(Transaction.GetObject(DBObject.Id, OpenMode.ForWrite), Entity)
                'Entity.UpgradeOpen()
                Entity.Erase()
                Entity.Dispose()
                Return True
            Catch
                Return False
            End Try
        End Function


        ''' <summary>
        ''' Move o conjunto de seleção
        ''' </summary>
        ''' <param name="SelectionSet">Seleção</param>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <param name="PointDisplacement">Ponto com os valores de deslocamento</param>
        ''' <returns>ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function Move(SelectionSet As SelectionSet, PointReference As Point3d, PointDisplacement As Point3d) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Entity As Entity
            Dim Matrix3d As Matrix3d = Matrix3d.Displacement(PointReference.GetVectorTo(PointDisplacement))
            Try
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        For Each ObjectId As ObjectId In SelectionSet.GetObjectIds()
                            Entity = Transaction.GetObject(ObjectId, OpenMode.ForWrite, False)
                            Entity.UpgradeOpen()
                            Entity.TransformBy(Matrix3d)
                        Next
                        Transaction.Commit()
                    End Using
                End Using
                Return New ObjectIdCollection(SelectionSet.GetObjectIds)
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Move o conjunto de seleção
        ''' </summary>
        ''' <param name="ObjectIdCollection">Coleção de ID´s</param>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <param name="PointDisplacement">Ponto com os valores de deslocamento</param>
        ''' <returns>ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function Move(ObjectIdCollection As ObjectIdCollection, PointReference As Point3d, PointDisplacement As Point3d) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Entity As Entity
            Dim Matrix3d As Matrix3d = Matrix3d.Displacement(PointReference.GetVectorTo(PointDisplacement))
            Try
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        For Each ObjectId As ObjectId In ObjectIdCollection
                            Entity = Transaction.GetObject(ObjectId, OpenMode.ForWrite, False)
                            Entity.UpgradeOpen()
                            Entity.TransformBy(Matrix3d)
                        Next
                        Transaction.Commit()
                    End Using
                End Using
                Return ObjectIdCollection
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Move o conjunto de seleção
        ''' </summary>
        ''' <param name="DBObjectCollection">Coleção de itens do banco de dados</param>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <param name="PointDisplacement">Ponto com os valores de deslocamento</param>
        ''' <returns>ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function Move(DBObjectCollection As DBObjectCollection, PointReference As Point3d, PointDisplacement As Point3d) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Entity As Entity
            Dim Matrix3d As Matrix3d = Matrix3d.Displacement(PointReference.GetVectorTo(PointDisplacement))
            Try
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        For Each DBObject As DBObject In DBObjectCollection
                            Entity = Transaction.GetObject(DBObject.ObjectId, OpenMode.ForWrite, False)
                            Entity.UpgradeOpen()
                            Entity.TransformBy(Matrix3d)
                        Next
                        Transaction.Commit()
                    End Using
                End Using
                Return DBObjectCollection
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Move entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <param name="PointDisplacement">Ponto de deslocamento</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Move(Entity As Entity, PointReference As Point3d, PointDisplacement As Point3d) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Matrix3d As Matrix3d = Matrix3d.Displacement(PointReference.GetVectorTo(PointDisplacement))
            Try
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite, False)
                        Entity.UpgradeOpen()
                        Entity.TransformBy(Matrix3d)
                        Transaction.Commit()
                    End Using
                End Using
                Return Entity
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Move entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <param name="PointDisplacement">Ponto de deslocamento</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Move(Transaction As Transaction, Entity As Entity, PointReference As Point3d, PointDisplacement As Point3d) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Matrix3d As Matrix3d = Matrix3d.Displacement(PointReference.GetVectorTo(PointDisplacement))
            Try
                Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite, False)
                Entity.UpgradeOpen()
                Entity.TransformBy(Matrix3d)
                Return Entity
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Rotaciona a entidade
        ''' </summary>
        ''' <param name="Entity">Entidade a ser totacionada</param>
        ''' <param name="BasePoint">Ponto base para rotação</param>
        ''' <param name="Angle">Ângulo (Degree)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Rotate(Entity As Entity, BasePoint As Point3d, Angle As Double) As Object
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            Dim Matrix3d As Matrix3d
            Angle = Engine2.Geometry.DegreeToRadian(Angle)
            Try
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                        BlockTableRecord = Transaction.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite, False)
                        Entity.UpgradeOpen()
                        Matrix3d = Document.Editor.CurrentUserCoordinateSystem
                        Entity.TransformBy(Matrix3d.Rotation(Angle, Matrix3d.CoordinateSystem3d.Zaxis, BasePoint))
                        Transaction.Commit()
                    End Using
                End Using
                Return Entity
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Rotaciona a entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Entity">Entidade a ser totacionada</param>
        ''' <param name="BasePoint">Ponto base para rotação</param>
        ''' <param name="Angle">Ângulo (Degree)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Rotate(Transaction As Transaction, Entity As Entity, BasePoint As Point3d, Angle As Double) As Object
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            Dim Matrix3d As Matrix3d
            Angle = Engine2.Geometry.DegreeToRadian(Angle)
            Try
                BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                BlockTableRecord = Transaction.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite, False)
                Entity.UpgradeOpen()
                Matrix3d = Document.Editor.CurrentUserCoordinateSystem
                Entity.TransformBy(Matrix3d.Rotation(Angle, Matrix3d.CoordinateSystem3d.Zaxis, BasePoint))
                Return Entity
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Cria cantos arredondados entre segmentos de linha
        ''' </summary>
        ''' <param name="Line1">Linha 1</param>
        ''' <param name="Line2">Linha 2</param>
        ''' <param name="Radius">Raio</param>
        ''' <returns>DBObjectCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function Fillet(Line1 As Line, Line2 As Line, Radius As Double) As DBObjectCollection
            Fillet = New DBObjectCollection
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            If Radius < Line1.Length And Radius < Line2.Length Then
                Dim Intersection As Point3d
                Dim PerpendicularInfor1 As PerpendicularInfor = Nothing
                Dim PerpendicularInfor2 As PerpendicularInfor = Nothing
                Dim PerpendicularInfor3 As PerpendicularInfor = Nothing
                Dim PerpendicularInfor4 As PerpendicularInfor = Nothing
                Dim Mediatriz As Point3d
                Dim Arc As Arc
                Dim VirtualParallelPoints As New List(Of Point3dCollection)
                Dim ParallelPoints As List(Of Point3dCollection)
                Dim Point3dCollection As Point3dCollection
                ParallelPoints = Engine2.Geometry.VirtualParallelPoints(Line1.StartPoint, Line1.EndPoint, Radius)
                For Each Point3dCollection In ParallelPoints
                    VirtualParallelPoints.Add(Point3dCollection)
                Next
                ParallelPoints = Engine2.Geometry.VirtualParallelPoints(Line2.StartPoint, Line2.EndPoint, Radius)
                For Each Point3dCollection In ParallelPoints
                    VirtualParallelPoints.Add(Point3dCollection)
                Next
                Point3dCollection = Engine2.Geometry.VirtualIntersection(VirtualParallelPoints, , Intersect.ExtendBoth)
                If Point3dCollection.Count = 4 Then
                    For Each Point3d As Point3d In Point3dCollection
                        PerpendicularInfor1 = Engine2.Geometry.GetPerpendicular(Line1, Point3d)
                        PerpendicularInfor2 = Engine2.Geometry.GetPerpendicular(Line2, Point3d)
                        If PerpendicularInfor1.VirtualPerpendicular = False And PerpendicularInfor2.VirtualPerpendicular = False Then
                            Intersection = PerpendicularInfor1.Origin
                            Exit For
                        End If
                    Next
                    For Each Point3d As Point3d In Point3dCollection
                        PerpendicularInfor3 = Engine2.Geometry.GetPerpendicular(Line1, Point3d)
                        PerpendicularInfor4 = Engine2.Geometry.GetPerpendicular(Line2, Point3d)
                        If PerpendicularInfor3.VirtualPerpendicular = True And PerpendicularInfor4.VirtualPerpendicular = True Then
                            Mediatriz = Engine2.Geometry.PolarPoint3d(Intersection, Engine2.Geometry.GetAngle(Intersection, Point3d, Geometry.eAngleFormat.Degrees), Radius)
                            Exit For
                        End If
                    Next
                    If PerpendicularInfor1.VirtualPerpendicular = False And PerpendicularInfor2.VirtualPerpendicular = False Then
                        Using DocumentLock As DocumentLock = Document.LockDocument
                            Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                                Line1 = Transaction.GetObject(Line1.ObjectId, OpenMode.ForRead)
                                Line2 = Transaction.GetObject(Line2.ObjectId, OpenMode.ForRead)
                                Line1.UpgradeOpen()
                                If Line1.StartPoint.DistanceTo(PerpendicularInfor1.PerpendicularPoint) < Line1.EndPoint.DistanceTo(PerpendicularInfor1.PerpendicularPoint) Then
                                    Line1.StartPoint = PerpendicularInfor1.PerpendicularPoint
                                Else
                                    Line1.EndPoint = PerpendicularInfor1.PerpendicularPoint
                                End If
                                Line2.UpgradeOpen()
                                If Line2.StartPoint.DistanceTo(PerpendicularInfor2.PerpendicularPoint) < Line2.EndPoint.DistanceTo(PerpendicularInfor2.PerpendicularPoint) Then
                                    Line2.StartPoint = PerpendicularInfor2.PerpendicularPoint
                                Else
                                    Line2.EndPoint = PerpendicularInfor2.PerpendicularPoint
                                End If
                                Arc = Engine2.Drawing.DrawArc(Transaction, PerpendicularInfor1.PerpendicularPoint, Mediatriz, PerpendicularInfor2.PerpendicularPoint)
                                Transaction.Commit()
                                Fillet.Add(Line1)
                                Fillet.Add(Arc)
                                Fillet.Add(Line2)
                            End Using
                        End Using
                    Else
                        Editor.WriteMessage(vbLf & "Impossível determinar curva para intersecção entre linhas." & vbLf)
                        Fillet.Clear()
                    End If
                Else
                    Editor.WriteMessage(vbLf & "Impossível determinar curva para intersecção entre linhas." & vbLf)
                    Fillet.Clear()
                End If
            Else
                Editor.WriteMessage(vbLf & "Impossível criar a curva, o raio é igual ou maior que o comprimento de um dos segmentos." & vbLf)
                Fillet.Clear()
            End If
            Return Fillet
        End Function

        ''' <summary>
        ''' Cria cantos arredondados entre segmentos de linha
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Line1">Linha 1</param>
        ''' <param name="Line2">Linha 2</param>
        ''' <param name="Radius">Raio</param>
        ''' <returns>DBObjectCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function Fillet(Transaction As Transaction, Line1 As Line, Line2 As Line, Radius As Double) As DBObjectCollection
            Fillet = New DBObjectCollection
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            If Radius < Line1.Length And Radius < Line2.Length Then
                Dim Intersection As Point3d
                Dim PerpendicularInfor1 As PerpendicularInfor = Nothing
                Dim PerpendicularInfor2 As PerpendicularInfor = Nothing
                Dim PerpendicularInfor3 As PerpendicularInfor = Nothing
                Dim PerpendicularInfor4 As PerpendicularInfor = Nothing
                Dim Mediatriz As Point3d
                Dim Arc As Arc
                Dim VirtualParallelPoints As New List(Of Point3dCollection)
                Dim ParallelPoints As List(Of Point3dCollection)
                Dim Point3dCollection As Point3dCollection
                ParallelPoints = Engine2.Geometry.VirtualParallelPoints(Line1.StartPoint, Line1.EndPoint, Radius)
                For Each Point3dCollection In ParallelPoints
                    VirtualParallelPoints.Add(Point3dCollection)
                Next
                ParallelPoints = Engine2.Geometry.VirtualParallelPoints(Line2.StartPoint, Line2.EndPoint, Radius)
                For Each Point3dCollection In ParallelPoints
                    VirtualParallelPoints.Add(Point3dCollection)
                Next
                Point3dCollection = Engine2.Geometry.VirtualIntersection(VirtualParallelPoints, , Intersect.ExtendBoth)
                If Point3dCollection.Count = 4 Then
                    For Each Point3d As Point3d In Point3dCollection
                        PerpendicularInfor1 = Engine2.Geometry.GetPerpendicular(Line1, Point3d)
                        PerpendicularInfor2 = Engine2.Geometry.GetPerpendicular(Line2, Point3d)
                        If PerpendicularInfor1.VirtualPerpendicular = False And PerpendicularInfor2.VirtualPerpendicular = False Then
                            Intersection = PerpendicularInfor1.Origin
                            Exit For
                        End If
                    Next
                    For Each Point3d As Point3d In Point3dCollection
                        PerpendicularInfor3 = Engine2.Geometry.GetPerpendicular(Line1, Point3d)
                        PerpendicularInfor4 = Engine2.Geometry.GetPerpendicular(Line2, Point3d)
                        If PerpendicularInfor3.VirtualPerpendicular = True And PerpendicularInfor4.VirtualPerpendicular = True Then
                            Mediatriz = Engine2.Geometry.PolarPoint3d(Intersection, Engine2.Geometry.GetAngle(Intersection, Point3d, Geometry.eAngleFormat.Degrees), Radius)
                            Exit For
                        End If
                    Next
                    If PerpendicularInfor1.VirtualPerpendicular = False And PerpendicularInfor2.VirtualPerpendicular = False Then
                        Line1 = Transaction.GetObject(Line1.ObjectId, OpenMode.ForRead)
                        Line2 = Transaction.GetObject(Line2.ObjectId, OpenMode.ForRead)
                        Line1.UpgradeOpen()
                        If Line1.StartPoint.DistanceTo(PerpendicularInfor1.PerpendicularPoint) < Line1.EndPoint.DistanceTo(PerpendicularInfor1.PerpendicularPoint) Then
                            Line1.StartPoint = PerpendicularInfor1.PerpendicularPoint
                        Else
                            Line1.EndPoint = PerpendicularInfor1.PerpendicularPoint
                        End If
                        Line2.UpgradeOpen()
                        If Line2.StartPoint.DistanceTo(PerpendicularInfor2.PerpendicularPoint) < Line2.EndPoint.DistanceTo(PerpendicularInfor2.PerpendicularPoint) Then
                            Line2.StartPoint = PerpendicularInfor2.PerpendicularPoint
                        Else
                            Line2.EndPoint = PerpendicularInfor2.PerpendicularPoint
                        End If
                        Arc = Engine2.Drawing.DrawArc(Transaction, PerpendicularInfor1.PerpendicularPoint, Mediatriz, PerpendicularInfor2.PerpendicularPoint)
                        Transaction.Commit()
                        Fillet.Add(Line1)
                        Fillet.Add(Arc)
                        Fillet.Add(Line2)
                    Else
                        Editor.WriteMessage(vbLf & "Impossível determinar curva para intersecção entre linhas." & vbLf)
                        Fillet.Clear()
                    End If
                Else
                    Editor.WriteMessage(vbLf & "Impossível determinar curva para intersecção entre linhas." & vbLf)
                    Fillet.Clear()
                End If
            Else
                Editor.WriteMessage(vbLf & "Impossível criar a curva, o raio é igual ou maior que o comprimento de um dos segmentos." & vbLf)
                Fillet.Clear()
            End If
            Return Fillet
        End Function

        ''' <summary>
        ''' Arredonda os vértices de uma polylinha
        ''' </summary>
        ''' <param name="Polyline">Polyline</param>
        ''' <param name="Radius">Raio</param>
        ''' <param name="LineWeight">Espessura</param>
        ''' <remarks></remarks>
        Public Shared Function Fillet(Polyline As Polyline, Radius As Double, Optional LineWeight As Double = 0) As Polyline
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim VertexPoint As Point2d
            Dim CircularArc2d As CircularArc2d
            Dim Index As Integer
            Dim Arc As Arc
            Dim PreviousSegment As Integer
            Dim Segment1 As LineSegment2d
            Dim Segment2 As LineSegment2d
            Dim Vector1 As Vector2d
            Dim Vector2 As Vector2d
            Dim Angle As Double
            Dim Distance As Double
            Dim FirstPoint As Point2d
            Dim NextPoint As Point2d
            Dim Bulge As Double
            Try
                Using DocumentLock As DocumentLock = Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        Polyline = Transaction.GetObject(Polyline.ObjectId, OpenMode.ForWrite)
                        Index = 0
                        While Index < Polyline.NumberOfVertices - 1
                            If Radius = 0 Then
                                Select Case Polyline.GetSegmentType(Index)
                                    Case SegmentType.Arc
                                        CircularArc2d = Polyline.GetArcSegment2dAt(Index)
                                        Arc = Engine2.EntityInteration.CircularArc2dToArc(CircularArc2d)
                                        VertexPoint = Engine2.Geometry.VertexOfArcTo2dPoint(Arc)
                                        Polyline.RemoveVertexAt(Index)
                                        Polyline.SetPointAt(Index, VertexPoint)
                                    Case Else
                                        Polyline.SetStartWidthAt(Index, LineWeight)
                                        Polyline.SetEndWidthAt(Index, LineWeight)
                                        Index = (Index + 1)
                                End Select
                            Else
                                If Index > 0 Then
                                    PreviousSegment = (Index - 1)
                                    If Polyline.GetSegmentType(PreviousSegment) = SegmentType.Line OrElse Polyline.GetSegmentType(Index) = SegmentType.Line Then
                                        Segment1 = Polyline.GetLineSegment2dAt(PreviousSegment)
                                        Segment2 = Polyline.GetLineSegment2dAt(Index)
                                        Vector1 = Segment1.StartPoint - Segment1.EndPoint
                                        Vector2 = Segment2.EndPoint - Segment2.StartPoint
                                        Angle = (Math.PI - Vector1.GetAngleTo(Vector2)) / 2.0
                                        Distance = Radius * Math.Tan(Angle)
                                        If Distance < Segment1.Length OrElse Distance < Segment2.Length Then
                                            FirstPoint = Segment1.EndPoint + Vector1.GetNormal * Distance
                                            NextPoint = Segment2.StartPoint + Vector2.GetNormal * Distance
                                            Bulge = Math.Tan(Angle / 2.0)
                                            If (((Segment1.EndPoint.X - Segment1.StartPoint.X) * (Segment2.EndPoint.Y - Segment1.StartPoint.Y) - (Segment1.EndPoint.Y - Segment1.StartPoint.Y) * (Segment2.EndPoint.X - Segment1.StartPoint.X)) < Tolerance.Global.EqualPoint) = True Then
                                                Bulge = -Bulge
                                            End If
                                            Polyline.AddVertexAt(Index, FirstPoint, Bulge, 0.0, 0.0)
                                            Polyline.SetPointAt(Index + 1, NextPoint)
                                            Polyline.SetStartWidthAt(Index, LineWeight)
                                            Polyline.SetEndWidthAt(Index, LineWeight)
                                            Polyline.SetStartWidthAt(Index + 1, LineWeight)
                                            Polyline.SetEndWidthAt(Index + 1, LineWeight)
                                            Index = (Index + 1)
                                        Else
                                            Editor.WriteMessage(vbLf & "Impossível criar a curva, o raio é igual ou maior que o comprimento de um dos segmentos." & vbLf)
                                        End If
                                    Else
                                        Editor.WriteMessage(vbLf & "Impossível determinar curva para intersecção entre segmentos." & vbLf)
                                    End If
                                End If
                                Polyline.SetStartWidthAt(Index, LineWeight)
                                Polyline.SetEndWidthAt(Index, LineWeight)
                                Index = (Index + 1)
                            End If
                        End While
                        Transaction.Commit()
                    End Using
                End Using
                Return Polyline
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Arredonda os vértices de uma polylinha
        ''' </summary>
        ''' <param name="Transaction">Transaction</param>
        ''' <param name="Polyline">Polyline</param>
        ''' <param name="Radius">Raio</param>
        ''' <param name="LineWeight">Espessura</param>
        ''' <remarks></remarks>
        Public Shared Function Fillet(Transaction As Transaction, Polyline As Polyline, Radius As Double, Optional LineWeight As Double = 0) As Polyline
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim VertexPoint As Point2d
            Dim CircularArc2d As CircularArc2d
            Dim Index As Integer
            Dim Arc As Arc
            Dim PreviousSegment As Integer
            Dim Segment1 As LineSegment2d
            Dim Segment2 As LineSegment2d
            Dim Vector1 As Vector2d
            Dim Vector2 As Vector2d
            Dim Angle As Double
            Dim Distance As Double
            Dim FirstPoint As Point2d
            Dim NextPoint As Point2d
            Dim Bulge As Double
            Try
                Polyline = Transaction.GetObject(Polyline.ObjectId, OpenMode.ForWrite)
                Index = 0
                While Index < Polyline.NumberOfVertices - 1
                    If Radius = 0 Then
                        Select Case Polyline.GetSegmentType(Index)
                            Case SegmentType.Arc
                                CircularArc2d = Polyline.GetArcSegment2dAt(Index)
                                Arc = Engine2.EntityInteration.CircularArc2dToArc(CircularArc2d)
                                VertexPoint = Engine2.Geometry.VertexOfArcTo2dPoint(Arc)
                                Polyline.RemoveVertexAt(Index)
                                Polyline.SetPointAt(Index, VertexPoint)
                            Case Else
                                Polyline.SetStartWidthAt(Index, LineWeight)
                                Polyline.SetEndWidthAt(Index, LineWeight)
                                Index = (Index + 1)
                        End Select
                    Else
                        If Index > 0 Then
                            PreviousSegment = (Index - 1)
                            If Polyline.GetSegmentType(PreviousSegment) = SegmentType.Line OrElse Polyline.GetSegmentType(Index) = SegmentType.Line Then
                                Segment1 = Polyline.GetLineSegment2dAt(PreviousSegment)
                                Segment2 = Polyline.GetLineSegment2dAt(Index)
                                Vector1 = Segment1.StartPoint - Segment1.EndPoint
                                Vector2 = Segment2.EndPoint - Segment2.StartPoint
                                Angle = (Math.PI - Vector1.GetAngleTo(Vector2)) / 2.0
                                Distance = Radius * Math.Tan(Angle)
                                If Distance < Segment1.Length OrElse Distance < Segment2.Length Then
                                    FirstPoint = Segment1.EndPoint + Vector1.GetNormal * Distance
                                    NextPoint = Segment2.StartPoint + Vector2.GetNormal * Distance
                                    Bulge = Math.Tan(Angle / 2.0)
                                    If (((Segment1.EndPoint.X - Segment1.StartPoint.X) * (Segment2.EndPoint.Y - Segment1.StartPoint.Y) - (Segment1.EndPoint.Y - Segment1.StartPoint.Y) * (Segment2.EndPoint.X - Segment1.StartPoint.X)) < Tolerance.Global.EqualPoint) = True Then
                                        Bulge = -Bulge
                                    End If
                                    Polyline.AddVertexAt(Index, FirstPoint, Bulge, 0.0, 0.0)
                                    Polyline.SetPointAt(Index + 1, NextPoint)
                                    Polyline.SetStartWidthAt(Index, LineWeight)
                                    Polyline.SetEndWidthAt(Index, LineWeight)
                                    Polyline.SetStartWidthAt(Index + 1, LineWeight)
                                    Polyline.SetEndWidthAt(Index + 1, LineWeight)
                                    Index = (Index + 1)
                                Else
                                    Editor.WriteMessage(vbLf & "Impossível criar a curva, o raio é igual ou maior que o comprimento de um dos segmentos." & vbLf)
                                End If
                            Else
                                Editor.WriteMessage(vbLf & "Impossível determinar curva para intersecção entre segmentos." & vbLf)
                            End If
                        End If
                        Polyline.SetStartWidthAt(Index, LineWeight)
                        Polyline.SetEndWidthAt(Index, LineWeight)
                        Index = (Index + 1)
                    End If
                End While
                Return Polyline
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Aplica fillet em um vértice específico
        ''' </summary>
        ''' <param name="Polyline">Polyline</param>
        ''' <param name="VerticeIndex">Vértice</param>
        ''' <param name="Radius">Raio</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function FilletVertex(Polyline As Polyline, VerticeIndex As Integer, Radius As Double) As Boolean
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim PreviousSegment As Integer
            Dim Segment1 As LineSegment2d
            Dim Segment2 As LineSegment2d
            Dim Vector1 As Vector2d
            Dim Vector2 As Vector2d
            Dim Angle As Double
            Dim Dist As Double
            Dim FirstPoint As Point2d
            Dim NextPoint As Point2d
            Dim Bulge As Double
            Try
                PreviousSegment = If(VerticeIndex = 0 AndAlso Polyline.Closed, Polyline.NumberOfVertices - 1, VerticeIndex - 1)
                If Polyline.GetSegmentType(PreviousSegment) = SegmentType.Line OrElse Polyline.GetSegmentType(VerticeIndex) = SegmentType.Line Then
                    Segment1 = Polyline.GetLineSegment2dAt(PreviousSegment)
                    Segment2 = Polyline.GetLineSegment2dAt(VerticeIndex)
                    Vector1 = Segment1.StartPoint - Segment1.EndPoint
                    Vector2 = Segment2.EndPoint - Segment2.StartPoint
                    Angle = (Math.PI - Vector1.GetAngleTo(Vector2)) / 2.0
                    Dist = Radius * Math.Tan(Angle)
                    If Dist < Segment1.Length OrElse Dist < Segment2.Length Then
                        FirstPoint = Segment1.EndPoint + Vector1.GetNormal() * Dist
                        NextPoint = Segment2.StartPoint + Vector2.GetNormal() * Dist
                        Bulge = Math.Tan(Angle / 2.0)
                        If (((Segment1.EndPoint.X - Segment1.StartPoint.X) * (Segment2.EndPoint.Y - Segment1.StartPoint.Y) - (Segment1.EndPoint.Y - Segment1.StartPoint.Y) * (Segment2.EndPoint.X - Segment1.StartPoint.X)) < Tolerance.Global.EqualPoint) = True Then
                            Bulge = -Bulge
                        End If
                        Using DocumentLock As DocumentLock = Document.LockDocument
                            Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                                Polyline = Transaction.GetObject(Polyline.ObjectId, OpenMode.ForWrite)
                                Polyline.AddVertexAt(VerticeIndex, FirstPoint, Bulge, 0.0, 0.0)
                                Polyline.SetPointAt(VerticeIndex + 1, NextPoint)
                                Transaction.Commit()
                            End Using
                        End Using
                        Return True
                    End If
                End If
                Return False
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Aplica fillet em um vértice específico
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Polyline">Polyline</param>
        ''' <param name="VerticeIndex">Vértice</param>
        ''' <param name="Radius">Raio</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function FilletVertex(Transaction As Transaction, Polyline As Polyline, VerticeIndex As Integer, Radius As Double) As Boolean
            Dim PreviousSegment As Integer
            Dim Segment1 As LineSegment2d
            Dim Segment2 As LineSegment2d
            Dim Vector1 As Vector2d
            Dim Vector2 As Vector2d
            Dim Angle As Double
            Dim Dist As Double
            Dim FirstPoint As Point2d
            Dim NextPoint As Point2d
            Dim Bulge As Double
            Try
                PreviousSegment = If(VerticeIndex = 0 AndAlso Polyline.Closed, Polyline.NumberOfVertices - 1, VerticeIndex - 1)
                If Polyline.GetSegmentType(PreviousSegment) = SegmentType.Line OrElse Polyline.GetSegmentType(VerticeIndex) = SegmentType.Line Then
                    Segment1 = Polyline.GetLineSegment2dAt(PreviousSegment)
                    Segment2 = Polyline.GetLineSegment2dAt(VerticeIndex)
                    Vector1 = Segment1.StartPoint - Segment1.EndPoint
                    Vector2 = Segment2.EndPoint - Segment2.StartPoint
                    Angle = (Math.PI - Vector1.GetAngleTo(Vector2)) / 2.0
                    Dist = Radius * Math.Tan(Angle)
                    If Dist < Segment1.Length OrElse Dist < Segment2.Length Then
                        FirstPoint = Segment1.EndPoint + Vector1.GetNormal() * Dist
                        NextPoint = Segment2.StartPoint + Vector2.GetNormal() * Dist
                        Bulge = Math.Tan(Angle / 2.0)
                        If (((Segment1.EndPoint.X - Segment1.StartPoint.X) * (Segment2.EndPoint.Y - Segment1.StartPoint.Y) - (Segment1.EndPoint.Y - Segment1.StartPoint.Y) * (Segment2.EndPoint.X - Segment1.StartPoint.X)) < Tolerance.Global.EqualPoint) = True Then
                            Bulge = -Bulge
                        End If
                        Polyline = Transaction.GetObject(Polyline.ObjectId, OpenMode.ForWrite)
                        Polyline.AddVertexAt(VerticeIndex, FirstPoint, Bulge, 0.0, 0.0)
                        Polyline.SetPointAt(VerticeIndex + 1, NextPoint)
                        Return True
                    End If
                End If
                Return False
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Remove as curvas de uma polyline
        ''' </summary>
        ''' <param name="Transaction"></param>
        ''' <param name="Polyline"></param>
        ''' <param name="Closed"></param>
        ''' <param name="LineWeight"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Decurve(Transaction As Transaction, Polyline As Polyline, Optional Closed As Boolean = False, Optional LineWeight As Double = 0) As Boolean
            Dim Point2dCollection As New Point2dCollection
            Dim Point2d As Point2d
            Try
                For Vertex As Integer = 0 To Polyline.NumberOfVertices - 1
                    Point2dCollection.Add(Polyline.GetPoint2dAt(Vertex))
                Next
                While Polyline.NumberOfVertices <> 1
                    Polyline.RemoveVertexAt(0)
                End While
                Polyline = Transaction.GetObject(Polyline.ObjectId, OpenMode.ForWrite)
                With Polyline
                    For Index = 0 To Point2dCollection.Count - 1
                        Point2d = Point2dCollection.Item(Index)
                        If Index = 0 Then
                            .SetPointAt(Index, Point2d)
                        Else
                            .AddVertexAt(Index, Point2d, 0, 0, 0)
                        End If
                        .SetStartWidthAt(Index, LineWeight)
                        .SetEndWidthAt(Index, LineWeight)
                    Next
                    .Closed = Closed
                End With
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Fecha uma polyline
        ''' </summary>
        ''' <param name="Polyline">Polyline</param>
        ''' <remarks></remarks>
        Public Shared Function Close(Polyline As Polyline) As Polyline
            Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                Using Transaction As Transaction = acadDocument.Database.TransactionManager.StartTransaction()
                    Polyline = Transaction.GetObject(Polyline.ObjectId, OpenMode.ForRead)
                    If Polyline.Closed = False Then
                        Polyline.UpgradeOpen()
                        Polyline.Closed = True
                        Transaction.Commit()
                    End If
                    Transaction.Dispose()
                End Using
            End Using
            Return Polyline
        End Function

        ''' <summary>
        ''' Abre uma polyline
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Polyline">Polyline</param>
        ''' <remarks></remarks>
        Public Shared Function Close(Transaction As Transaction, Polyline As Polyline) As Polyline
            Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            If Polyline.Closed = False Then
                Polyline.Closed = True
            End If
            Return Polyline
        End Function

        ''' <summary>
        ''' Fecha uma polyline
        ''' </summary>
        ''' <param name="Polyline">Polyline</param>
        ''' <remarks></remarks>
        Public Shared Function Open(Polyline As Polyline) As Polyline
            Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                Using Transaction As Transaction = acadDocument.Database.TransactionManager.StartTransaction()
                    Polyline = Transaction.GetObject(Polyline.ObjectId, OpenMode.ForRead)
                    If Polyline.Closed = True Then
                        Polyline.UpgradeOpen()
                        Polyline.Closed = False
                        Transaction.Commit()
                    End If
                    Transaction.Dispose()
                End Using
            End Using
            Return Polyline
        End Function

        ''' <summary>
        ''' Abre uma polyline
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Polyline">Polyline</param>
        ''' <remarks></remarks>
        Public Shared Function Open(Transaction As Transaction, Polyline As Polyline) As Polyline
            Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            If Polyline.Closed = True Then
                Polyline.Closed = False
            End If
            Return Polyline
        End Function

        ''' <summary>
        ''' Atualiza bloco
        ''' </summary>
        ''' <param name="BlockReference">Bloco</param>
        ''' <param name="Rotation">Rotação (Radian)</param>
        ''' <param name="Scale3d">Escala</param>
        ''' <param name="Position">Posição</param>
        ''' <param name="AngleFormat">Formato informado para a rotação</param>
        ''' <remarks></remarks>
        Public Shared Function ChangeBlock(BlockReference As BlockReference, Rotation As Double, Scale3d As Scale3d, Position As Point3d, AngleFormat As Engine2.Geometry.eAngleFormat) As Boolean
            ChangeBlock = False
            Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                Using Transaction As Transaction = acadDocument.Database.TransactionManager.StartTransaction()
                    BlockReference = Transaction.GetObject(BlockReference.ObjectId, OpenMode.ForWrite)
                    If BlockReference.Position.DistanceTo(Position) > Tolerance.Global.EqualPoint Or Position.Z <> BlockReference.Position.Z Then
                        BlockReference.Position = Position
                        ChangeBlock = True
                    End If
                    If AngleFormat = Geometry.eAngleFormat.Degrees Then
                        If Engine2.Geometry.RadianToDegree(BlockReference.Rotation) <> Rotation Then
                            BlockReference.Rotation = Engine2.Geometry.DegreeToRadian(Rotation)
                            ChangeBlock = True
                        End If
                    Else
                        If BlockReference.Rotation <> Rotation Then
                            BlockReference.Rotation = Rotation
                            ChangeBlock = True
                        End If
                    End If
                    If BlockReference.ScaleFactors.Equals(Scale3d) = False Then
                        BlockReference.ScaleFactors = Scale3d
                        ChangeBlock = True
                    End If
                    Transaction.Commit()
                End Using
            End Using
            Return ChangeBlock
        End Function

        ''' <summary>
        ''' Atualiza bloco
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="BlockReference">Bloco</param>
        ''' <param name="Rotation">Rotação (Radian)</param>
        ''' <param name="Scale3d">Escala</param>
        ''' <param name="Position">Posição</param>
        ''' <param name="AngleFormat">Formato informado para a rotação</param>
        ''' <remarks></remarks>
        Public Shared Function ChangeBlock(Transaction As Transaction, BlockReference As BlockReference, Rotation As Double, Scale3d As Scale3d, Position As Point3d, AngleFormat As Engine2.Geometry.eAngleFormat) As Boolean
            ChangeBlock = False
            BlockReference = Transaction.GetObject(BlockReference.ObjectId, OpenMode.ForWrite)
            If BlockReference.Position.DistanceTo(Position) > Tolerance.Global.EqualPoint Or Position.Z <> BlockReference.Position.Z Then
                BlockReference.Position = Position
                ChangeBlock = True
            End If
            If AngleFormat = Geometry.eAngleFormat.Degrees Then
                If Engine2.Geometry.RadianToDegree(BlockReference.Rotation) <> Rotation Then
                    BlockReference.Rotation = Engine2.Geometry.DegreeToRadian(Rotation)
                    ChangeBlock = True
                End If
            Else
                If BlockReference.Rotation <> Rotation Then
                    BlockReference.Rotation = Rotation
                    ChangeBlock = True
                End If
            End If
            If BlockReference.ScaleFactors.Equals(Scale3d) = False Then
                BlockReference.ScaleFactors = Scale3d
                ChangeBlock = True
            End If
            Return ChangeBlock
        End Function

        ''' <summary>
        ''' Aplica a espessura na polyline
        ''' </summary>
        ''' <param name="Polyline">Polyline</param>
        ''' <param name="LineWeight">Espessura</param>
        ''' <returns>Polyline</returns>
        ''' <remarks></remarks>
        Public Shared Function SetLineWeight(Polyline As Polyline, LineWeight As Double) As Polyline
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Index As Integer
            Try
                Using DocumentLock As DocumentLock = Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        Polyline = Transaction.GetObject(Polyline.ObjectId, OpenMode.ForWrite)
                        Index = 0
                        While Index < Polyline.NumberOfVertices - 1
                            If Polyline.GetSegmentType(Index) = SegmentType.Arc Or Polyline.GetSegmentType(Index) = SegmentType.Line Then
                                Polyline.SetStartWidthAt(Index, LineWeight)
                                Polyline.SetEndWidthAt(Index, LineWeight)
                                Polyline.SetStartWidthAt(Index + 1, LineWeight)
                                Polyline.SetEndWidthAt(Index + 1, LineWeight)
                            End If
                            Index = (Index + 2)
                        End While
                        Transaction.Commit()
                    End Using
                End Using
                Return Polyline
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Aplica a espessura na polyline
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Polyline">Polyline</param>
        ''' <param name="LineWeight">Espessura</param>
        ''' <returns>Polyline</returns>
        ''' <remarks></remarks>
        Public Shared Function SetLineWeight(Transaction As Transaction, Polyline As Polyline, LineWeight As Double) As Polyline
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Index As Integer
            Try
                Polyline = Transaction.GetObject(Polyline.ObjectId, OpenMode.ForWrite)
                Index = 0
                While Index < Polyline.NumberOfVertices - 1
                    If Polyline.GetSegmentType(Index) = SegmentType.Arc Or Polyline.GetSegmentType(Index) = SegmentType.Line Then
                        Polyline.SetStartWidthAt(Index, LineWeight)
                        Polyline.SetEndWidthAt(Index, LineWeight)
                        Polyline.SetStartWidthAt(Index + 1, LineWeight)
                        Polyline.SetEndWidthAt(Index + 1, LineWeight)
                    End If
                    Index = (Index + 2)
                End While
                Return Polyline
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Edita DBText
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DBText">DBText</param>
        ''' <param name="TextString">Texto</param>
        ''' <param name="Position">Local da inserção</param>
        ''' <param name="Height">Altura do texto</param>
        ''' <param name="TextStyle">Estilo de texto</param>
        ''' <param name="Justify">Posicionamento</param>
        ''' <param name="Rotation">Rotação (radianos)</param>
        ''' <param name="IsMirroredInX">Espelhamento em X</param>
        ''' <param name="IsMirroredInY">Espelhamento em Y</param>
        ''' <param name="InMemory"> Setar este valor se o item estiver sendo processado em memória</param>
        ''' <returns>DBText</returns>
        ''' <remarks></remarks>
        Public Shared Function DBTextEdit(Transaction As Transaction, DBText As DBText, TextString As String, Optional Position As Object = Nothing, Optional Height As Object = Nothing, Optional TextStyle As String = "Standard", Optional Justify As Autodesk.AutoCAD.DatabaseServices.AttachmentPoint = AttachmentPoint.MiddleCenter, Optional Rotation As Double = 0, Optional IsMirroredInX As Object = Nothing, Optional IsMirroredInY As Object = Nothing, Optional InMemory As Boolean = False) As DBText
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim TextStyleTable As TextStyleTable
            Dim PreviousDataBase As Database
            Try
                TextStyleTable = Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForRead)
                DBText = Transaction.GetObject(DBText.ObjectId, OpenMode.ForWrite)
                With DBText
                    If InMemory = True Then
                        PreviousDataBase = HostApplicationServices.WorkingDatabase
                        HostApplicationServices.WorkingDatabase = Database
                        DBText.AdjustAlignment(Database)
                        HostApplicationServices.WorkingDatabase = PreviousDataBase
                    Else
                        DBText.AdjustAlignment(Database)
                    End If
                    If IsNothing(Position) = False Then
                        .AlignmentPoint = .Position
                    End If
                    .HorizontalMode = TextHorizontalMode.TextCenter
                    .VerticalMode = TextVerticalMode.TextVerticalMid
                    If IsNothing(Justify) = False Then
                        .Justify = Justify
                    End If
                    .Position = .AlignmentPoint
                    If IsNothing(TextStyle) = False Then
                        .TextStyleId = TextStyleTable(TextStyle)
                    End If
                    .TextString = TextString
                    If IsNothing(Height) = False Then
                        .Height = Height
                    End If
                    If IsNothing(Rotation) = False Then
                        .Rotation = Rotation
                    End If
                    If IsNothing(IsMirroredInX) = False Then
                        .IsMirroredInX = IsMirroredInX
                    End If
                    If IsNothing(IsMirroredInY) = False Then
                        .IsMirroredInY = IsMirroredInY
                    End If
                End With
                Return DBText
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Edita DBText
        ''' </summary>
        ''' <param name="DBText">DBText</param>
        ''' <param name="TextString">Texto</param>
        ''' <param name="Position">Local da inserção</param>
        ''' <param name="Height">Altura do texto</param>
        ''' <param name="TextStyle">Estilo de texto</param>
        ''' <param name="Justify">Posicionamento</param>
        ''' <param name="Rotation">Rotação (radianos)</param>
        ''' <param name="IsMirroredInX">Espelhamento em X</param>
        ''' <param name="IsMirroredInY">Espelhamento em Y</param>
        ''' <param name="InMemory"> Setar este valor se o item estiver sendo processado em memória</param>
        ''' <returns>DBText</returns>
        ''' <remarks></remarks>
        Public Shared Function DBTextEdit(DBText As DBText, TextString As String, Optional Position As Object = Nothing, Optional Height As Object = Nothing, Optional TextStyle As String = "Standard", Optional Justify As Autodesk.AutoCAD.DatabaseServices.AttachmentPoint = AttachmentPoint.MiddleCenter, Optional Rotation As Double = 0, Optional IsMirroredInX As Object = Nothing, Optional IsMirroredInY As Object = Nothing, Optional InMemory As Boolean = False) As DBText
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim TextStyleTable As TextStyleTable
            Dim PreviousDataBase As Database
            Try
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        TextStyleTable = Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForRead)
                        DBText = Transaction.GetObject(DBText.ObjectId, OpenMode.ForWrite)
                        With DBText
                            If InMemory = True Then
                                PreviousDataBase = HostApplicationServices.WorkingDatabase
                                HostApplicationServices.WorkingDatabase = Database
                                DBText.AdjustAlignment(Database)
                                HostApplicationServices.WorkingDatabase = PreviousDataBase
                            Else
                                DBText.AdjustAlignment(Database)
                            End If
                            If IsNothing(Position) = False Then
                                .AlignmentPoint = .Position
                            End If
                            .HorizontalMode = TextHorizontalMode.TextCenter
                            .VerticalMode = TextVerticalMode.TextVerticalMid
                            If IsNothing(Justify) = False Then
                                .Justify = Justify
                            End If
                            .Position = .AlignmentPoint
                            If IsNothing(TextStyle) = False Then
                                .TextStyleId = TextStyleTable(TextStyle)
                            End If
                            .TextString = TextString
                            If IsNothing(Height) = False Then
                                .Height = Height
                            End If
                            If IsNothing(Rotation) = False Then
                                .Rotation = Rotation
                            End If
                            If IsNothing(IsMirroredInX) = False Then
                                .IsMirroredInX = IsMirroredInX
                            End If
                            If IsNothing(IsMirroredInY) = False Then
                                .IsMirroredInY = IsMirroredInY
                            End If
                        End With
                        Transaction.Commit()
                    End Using
                End Using
                Return DBText
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Edita a escala de um tipo de linha
        ''' </summary>
        ''' <param name="Curve">Linha</param>
        ''' <param name="LinetypeScale">Escala</param>
        ''' <returns>Polyline</returns>
        ''' <remarks></remarks>
        Public Shared Function SetLinetypeScale(Curve As Curve, LinetypeScale As Double) As Curve
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Try
                Using DocumentLock As DocumentLock = Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        Curve = Transaction.GetObject(Curve.ObjectId, OpenMode.ForWrite)
                        Curve.LinetypeScale = LinetypeScale
                        Transaction.Commit()
                    End Using
                End Using
                Return Curve
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Edita a escala de um tipo de linha
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Curve">Linha</param>
        ''' <param name="LinetypeScale">Escala</param>
        ''' <returns>Polyline</returns>
        ''' <remarks></remarks>
        Public Shared Function SetLinetypeScale(Transaction As Transaction, Curve As Curve, LinetypeScale As Double) As Curve
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Try
                Curve = Transaction.GetObject(Curve.ObjectId, OpenMode.ForWrite)
                Curve.LinetypeScale = LinetypeScale
                Return Curve
            Catch
                Return Nothing
            End Try
        End Function


    End Class

End Namespace

