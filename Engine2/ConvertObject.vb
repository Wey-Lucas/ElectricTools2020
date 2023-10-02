'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.GraphicsInterface
Imports Autodesk.AutoCAD.Geometry

Namespace Engine2

    ''' <summary>
    ''' Classe para conversão de objetos nativos do AutoCAD
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class ConvertObject

        ''' <summary>
        ''' Retorna a entidade baseando-se no Handle
        ''' </summary>
        ''' <param name="HandleValue">HandleValue</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function HandleToEntity(ByVal HandleValue As Object) As Entity
            Try
                Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
                Dim acadEditor As Editor = acadDocument.Editor
                Dim acadDataBase As Database = acadDocument.Database
                Dim LongValue As Long = Convert.ToInt64(HandleValue, 16)
                Dim Handle As Handle = New Handle(LongValue)
                Dim ObjectId As ObjectId = acadDataBase.GetObjectId(False, Handle, 0)
                Using Transaction As Transaction = acadDataBase.TransactionManager.StartTransaction
                    HandleToEntity = Transaction.GetObject(ObjectId, OpenMode.ForRead)
                End Using
                Return HandleToEntity
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retorna a entidade baseando-se no Handle
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="HandleValue">HandleValue</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function HandleToEntity(Transaction As Transaction, ByVal HandleValue As Object) As Entity
            Try
                Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
                Dim acadEditor As Editor = acadDocument.Editor
                Dim acadDataBase As Database = acadDocument.Database
                Dim LongValue As Long = Convert.ToInt64(HandleValue, 16)
                Dim Handle As Handle = New Handle(LongValue)
                Dim ObjectId As ObjectId = acadDataBase.GetObjectId(False, Handle, 0)
                HandleToEntity = Transaction.GetObject(ObjectId, OpenMode.ForRead)
                Return HandleToEntity
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Converte objectID em entidade
        ''' </summary>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <param name="OpenMode">Modo de abertura</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ObjectIDToEntity(ByVal ObjectId As ObjectId, Optional OpenMode As OpenMode = OpenMode.ForRead) As Entity
            Try
                Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
                Dim acadEditor As Editor = acadDocument.Editor
                Dim acadDataBase As Database = acadDocument.Database
                Using Transaction As Transaction = acadDataBase.TransactionManager.StartTransaction
                    ObjectIDToEntity = DirectCast(Transaction.GetObject(ObjectId, OpenMode), Entity)
                End Using
                Return ObjectIDToEntity
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Converte objectID em entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <param name="OpenMode">Modo de abertura</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ObjectIDToEntity(Transaction As Transaction, ByVal ObjectId As ObjectId, Optional OpenMode As OpenMode = OpenMode.ForRead) As Entity
            Try
                Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
                Dim acadEditor As Editor = acadDocument.Editor
                Dim acadDataBase As Database = acadDocument.Database
                ObjectIDToEntity = DirectCast(Transaction.GetObject(ObjectId, OpenMode), Entity)
                Return ObjectIDToEntity
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Converte objectID em DBObject
        ''' </summary>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ObjectIDToDBObject(ByVal ObjectId As ObjectId) As DBObject
            Try
                Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
                Dim acadEditor As Editor = acadDocument.Editor
                Dim acadDataBase As Database = acadDocument.Database
                Using Transaction As Transaction = acadDataBase.TransactionManager.StartTransaction
                    ObjectIDToDBObject = DirectCast(Transaction.GetObject(ObjectId, OpenMode.ForRead), DBObject)
                End Using
                Return ObjectIDToDBObject
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Converte objectID em DBObject
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ObjectIDToDBObject(Transaction As Transaction, ByVal ObjectId As ObjectId) As DBObject
            Try
                Dim acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
                Dim acadEditor As Editor = acadDocument.Editor
                Dim acadDataBase As Database = acadDocument.Database
                ObjectIDToDBObject = DirectCast(Transaction.GetObject(ObjectId, OpenMode.ForRead), DBObject)
                Return ObjectIDToDBObject
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Converte DBObject em entidade
        ''' </summary>
        ''' <param name="DBObject">DBObject</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DBObjectToEntity(ByVal DBObject As DBObject) As Entity
            Try
                Return CType(DBObject, Entity)
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Converte ObjectIdCollection em List(Of ObjectId)
        ''' </summary>
        ''' <param name="ObjectIdCollection">ObjectIdCollection</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ObjectIdCollectionToListOfObjectId(ObjectIdCollection As ObjectIdCollection) As List(Of ObjectId)
            Try
                Dim ObjectIds(ObjectIdCollection.Count - 1) As ObjectId
                ObjectIdCollection.CopyTo(ObjectIds, 0)
                Return New List(Of ObjectId)(ObjectIds)
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Converte ObjectIdCollection em List(Of ObjectId)
        ''' </summary>
        ''' <param name="ListOfObjectId">List(Of ObjectId)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ListOfObjectIdToObjectIdCollection(ListOfObjectId As List(Of ObjectId)) As ObjectIdCollection
            Try
                Return New ObjectIdCollection(ListOfObjectId.ToArray)
            Catch
                Return Nothing
            End Try
        End Function

    End Class

End Namespace


