'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Namespace Engine2

    ''' <summary>
    ''' Explode entidades de desenho
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ExplodeObject

        ''' <summary>
        ''' Explode uma entidade
        ''' </summary>
        ''' <param name="Entity"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Explode(Entity As Entity) As Boolean
            Try
                Dim Database As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
                Dim Editor As Editor = Core.Application.DocumentManager.MdiActiveDocument.Editor
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        Dim DBObjectCollection As New DBObjectCollection()
                        Entity.UpgradeOpen()
                        Entity.Explode(DBObjectCollection)
                        Dim BlockTableRecord As BlockTableRecord = DirectCast(Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                        For Each DBObject As DBObject In DBObjectCollection
                            Entity = DirectCast(DBObject, Entity)
                            BlockTableRecord.AppendEntity(Entity)
                            Transaction.AddNewlyCreatedDBObject(Entity, True)
                        Next
                        Engine2.Edit.EraseEntity(Entity)
                        Transaction.Commit()
                    End Using
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

    End Class

End Namespace
