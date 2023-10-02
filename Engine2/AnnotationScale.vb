Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Internal

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2019
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    Public Class AnnotationScaleManager

        ''' <summary>
        ''' Seta a escala anotativa corrente na entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function SetCurrentAnnotationScale(Entity As Entity, Optional Transaction As Transaction = Nothing) As Boolean
            SetCurrentAnnotationScale = False
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim ObjectContextManager As ObjectContextManager
            Dim ObjectContextCollection As ObjectContextCollection
            If IsNothing(Transaction) = False Then
                Dim AnnotativeStates As AnnotativeStates = Entity.Annotative
                If AnnotativeStates = Autodesk.AutoCAD.DatabaseServices.AnnotativeStates.True Then
                    Entity = Transaction.GetObject(Entity.Id, OpenMode.ForWrite)
                    ObjectContextManager = Database.ObjectContextManager
                    ObjectContextCollection = ObjectContextManager.GetContextCollection("ACDB_ANNOTATIONSCALES")
                    ObjectContexts.AddContext(Entity, ObjectContextCollection.CurrentContext)
                    SetCurrentAnnotationScale = True
                End If
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Transaction = Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            SetCurrentAnnotationScale = SetCurrentAnnotationScale(Entity, Transaction)
                            Transaction.Commit()
                        Catch
                            SetCurrentAnnotationScale = False
                            Transaction.Abort()
                        End Try
                    End Using
                End Using
            End If
            Return SetCurrentAnnotationScale
        End Function

        ''' <summary>
        ''' Seta uma escala anotativa em uma entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="PaperUnits">Unidades do papel</param>
        ''' <param name="DrawingUnits">Unidades do desenho</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>ObjectContext</returns>
        ''' <remarks></remarks>
        Public Shared Function SetAnnotativeScale(ByVal Entity As Entity, ByVal PaperUnits As Double, DrawingUnits As Double, Optional Transaction As Transaction = Nothing) As ObjectContext
            SetAnnotativeScale = Nothing
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim AnnotationScale As AnnotationScale
            Dim NameScale As String
            Dim ObjectContextManager As ObjectContextManager
            Dim ObjectContextCollection As ObjectContextCollection
            If IsNothing(Transaction) = False Then
                NameScale = PaperUnits & ":" & DrawingUnits
                ObjectContextManager = Entity.Database.ObjectContextManager
                ObjectContextCollection = ObjectContextManager.GetContextCollection("ACDB_ANNOTATIONSCALES")
                If ObjectContextCollection.HasContext(NameScale) = False Then
                    AnnotationScale = New AnnotationScale()
                    AnnotationScale.Name = NameScale
                    AnnotationScale.PaperUnits = PaperUnits
                    AnnotationScale.DrawingUnits = DrawingUnits
                    ObjectContextCollection.AddContext(AnnotationScale)
                    SetAnnotativeScale = AnnotationScale
                Else
                    SetAnnotativeScale = ObjectContextCollection.GetContext(NameScale)
                End If
                Entity = Transaction.GetObject(Entity.Id, OpenMode.ForWrite)
                Entity.AddContext(SetAnnotativeScale)
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Transaction = Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            SetAnnotativeScale = SetAnnotativeScale(Entity, PaperUnits, DrawingUnits, Transaction)
                            Transaction.Commit()
                        Catch
                            SetAnnotativeScale = Nothing
                            Transaction.Abort()
                        End Try
                    End Using
                End Using
            End If
            Return SetAnnotativeScale
        End Function

        ''' <summary>
        ''' Obtem a coleção de escalas de anotação
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function GetAnnotationScales() As List(Of AnnotationScale)
            GetAnnotationScales = New List(Of AnnotationScale)
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim contextManager As ObjectContextManager = Database.ObjectContextManager
            Dim contextCollection As ObjectContextCollection
            Dim scale As AnnotationScale
            If IsNothing(contextManager) = False Then
                contextCollection = contextManager.GetContextCollection("ACDB_ANNOTATIONSCALES")
                If IsNothing(contextCollection) = False Then
                    For Each objCxt As ObjectContext In contextCollection
                        scale = TryCast(objCxt, AnnotationScale)
                        GetAnnotationScales.Add(scale)
                    Next
                End If
            End If
            GetAnnotationScales.Sort(Function(X As AnnotationScale, Y As AnnotationScale) X.Name.CompareTo(Y.Name))
            Return GetAnnotationScales
        End Function

    End Class

End Namespace


