Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Colors

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    Public Class ChangeColor

        ''' <summary>
        ''' Enumera as cores
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eColors
            Red = 1
            Yellow = 2
            Green = 3
            Cyan = 4
            Blue = 5
            Magenta = 6
            White = 7
        End Enum

        ''' <summary>
        ''' Muda a cor da entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="color">Cor</param>
        ''' <remarks></remarks>
        Public Shared Sub ChangeColor(Entity As Entity, color As eColors)
            Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acadDatabase As Database = acadDocument.Database
            Using acadDocument.LockDocument
                Using Transaction As Transaction = acadDocument.TransactionManager.StartTransaction()
                    If Entity.ColorIndex <> color Then
                        Dim LayerTable As LayerTable = TryCast(Transaction.GetObject(acadDatabase.LayerTableId, OpenMode.ForRead), LayerTable)
                        Entity = TryCast(Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite), Entity)
                        Entity.ColorIndex = color
                    End If
                    Transaction.Commit()
                    Transaction.Dispose()
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Muda a cor da entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="color">Cor</param>
        ''' <remarks></remarks>
        Public Shared Sub ChangeColor(Entity As Entity, color As Short)
            Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acadDatabase As Database = acadDocument.Database
            Using acadDocument.LockDocument
                Using Transaction As Transaction = acadDocument.TransactionManager.StartTransaction()
                    If Entity.ColorIndex <> color Then
                        Dim LayerTable As LayerTable = TryCast(Transaction.GetObject(acadDatabase.LayerTableId, OpenMode.ForRead), LayerTable)
                        Entity = TryCast(Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite), Entity)
                        Entity.ColorIndex = color
                    End If
                    Transaction.Commit()
                    Transaction.Dispose()
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Muda a cor da entidade
        ''' </summary>
        ''' <param name="ObjectIdCollection">Coleção de ID´s</param>
        ''' <param name="color">Cor</param>
        ''' <remarks></remarks>
        Public Shared Sub ChangeColor(ObjectIdCollection As ObjectIdCollection, color As eColors)
            Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acadDatabase As Database = acadDocument.Database
            Dim Entity As Entity
            Using acadDocument.LockDocument
                Using Transaction As Transaction = acadDocument.TransactionManager.StartTransaction()
                    For ini As Integer = 0 To ObjectIdCollection.Count - 1
                        Entity = Engine2.ConvertObject.ObjectIDToEntity(ObjectIdCollection.Item(ini))
                        If Entity.ColorIndex <> color Then
                            Dim LayerTable As LayerTable = TryCast(Transaction.GetObject(acadDatabase.LayerTableId, OpenMode.ForRead), LayerTable)
                            Entity = TryCast(Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite), Entity)
                            Entity.ColorIndex = color
                        End If
                    Next
                    Transaction.Commit()
                    Transaction.Dispose()
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Muda a cor da entidade
        ''' </summary>
        ''' <param name="ObjectIdCollection">Coleção de ID´s</param>
        ''' <param name="color">Cor</param>
        ''' <remarks></remarks>
        Public Shared Sub ChangeColor(ObjectIdCollection As ObjectIdCollection, color As Short)
            Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acadDatabase As Database = acadDocument.Database
            Dim Entity As Entity
            Using acadDocument.LockDocument
                Using Transaction As Transaction = acadDocument.TransactionManager.StartTransaction()
                    For ini As Integer = 0 To ObjectIdCollection.Count - 1
                        Entity = Engine2.ConvertObject.ObjectIDToEntity(ObjectIdCollection.Item(ini))
                        If Entity.ColorIndex <> color Then
                            Dim LayerTable As LayerTable = TryCast(Transaction.GetObject(acadDatabase.LayerTableId, OpenMode.ForRead), LayerTable)
                            Entity = TryCast(Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite), Entity)
                            Entity.ColorIndex = color
                        End If
                    Next
                    Transaction.Commit()
                    Transaction.Dispose()
                End Using
            End Using
        End Sub

    End Class

End Namespace