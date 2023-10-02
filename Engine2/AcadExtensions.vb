Imports System.Runtime.CompilerServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Extensões em entidades do AutoCad
    ''' </summary>
    ''' <remarks></remarks>
    Module AcadExtensions

#Region "Conversão de pontos UCS \ WCS"

        ''' <summary>
        ''' Retorna o ponto relativo a UCS
        ''' </summary>
        ''' <param name="Point3d"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToUCS(Point3d As Point3d) As Point3d
            Return Engine2.UCS.Point3dWcsToUcs(Point3d)
        End Function

        ''' <summary>
        ''' Retorna o ponto relativo a WCS
        ''' </summary>
        ''' <param name="Point3d"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToWCS(Point3d As Point3d) As Point3d
            Return Engine2.UCS.Point3dUcsToWcs(Point3d)
        End Function

#End Region

#Region "Conversão de objetos do AutoCad"

        ''' <summary>
        ''' Converte Handle em Entity
        ''' </summary>
        ''' <returns>Entity</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToEntity(Handle As Handle) As Entity
            Return Engine2.ConvertObject.HandleToEntity(Handle)
        End Function

        ''' <summary>
        ''' Converte Handle em Entity
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>Entity</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToEntity(Handle As Handle, Transaction As Transaction) As Entity
            Return Engine2.ConvertObject.HandleToEntity(Transaction, Handle)
        End Function

        ''' <summary>
        ''' Converte DBObject em Entity
        ''' </summary>
        ''' <returns>Entity</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToEntity(DBObject As DBObject) As Entity
            Return Engine2.ConvertObject.DBObjectToEntity(DBObject)
        End Function

        ''' <summary>
        ''' Converte ObjectId em Entity
        ''' </summary>
        ''' <returns>Entity</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToEntity(ObjectId As ObjectId) As Entity
            Return Engine2.ConvertObject.ObjectIDToEntity(ObjectId)
        End Function

        ''' <summary>
        ''' Converte ObjectId em Entity
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>Entity</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToEntity(ObjectId As ObjectId, Transaction As Transaction) As Entity
            Return Engine2.ConvertObject.ObjectIDToEntity(Transaction, ObjectId)
        End Function

        ''' <summary>
        ''' Converte ObjectId em DBObject
        ''' </summary>
        ''' <returns>DBObject</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToDBObject(ObjectId As ObjectId) As DBObject
            Return Engine2.ConvertObject.ObjectIDToDBObject(ObjectId)
        End Function

        ''' <summary>
        ''' Converte ObjectId em DBObject
        ''' </summary>
        ''' <param name="Transaction">Transaction</param>
        ''' <returns>DBObject</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToDBObject(ObjectId As ObjectId, Transaction As Transaction) As DBObject
            Return Engine2.ConvertObject.ObjectIDToDBObject(Transaction, ObjectId)
        End Function

#End Region



#Region "Métodos XData em DBObject"

        ''' <summary>
        ''' Verifica se existe um Xdata com chave específica
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ContainsXData(DBObject As DBObject, Key As String) As Boolean
            Return Engine2.XDataEngine.Contains(DBObject, Key)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(DBObject As DBObject, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(DBObject, Key, Value)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(DBObject As DBObject, Transaction As Transaction, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(Transaction, DBObject, Key, Value)
        End Function

        ''' <summary>
        ''' Lê XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function GetXData(DBObject As DBObject, Key As String) As Object
            Return Engine2.XDataEngine.GetXData(DBObject, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(DBObject As DBObject, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(DBObject, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(DBObject As DBObject, Transaction As Transaction, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(Transaction, DBObject, Key)
        End Function

#End Region

#Region "Métodos XData em ObjectId"

        ''' <summary>
        ''' Verifica se existe um Xdata com chave específica
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function XDataContains(ObjectId As ObjectId, Key As String) As Boolean
            Return Engine2.XDataEngine.Contains(ObjectId, Key)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(ObjectId As ObjectId, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(ObjectId, Key, Value)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(ObjectId As ObjectId, Transaction As Transaction, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(Transaction, ObjectId, Key, Value)
        End Function

        ''' <summary>
        ''' Lê XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function GetXData(ObjectId As ObjectId, Key As String) As Object
            Return Engine2.XDataEngine.GetXData(ObjectId, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(ObjectId As ObjectId, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(ObjectId, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(ObjectId As ObjectId, Transaction As Transaction, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(Transaction, ObjectId, Key)
        End Function

#End Region

#Region "Métodos XData em Polyline"

        ''' <summary>
        ''' Verifica se existe um Xdata com chave específica
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function XDataContains(Polyline As Polyline, Key As String) As Boolean
            Return Engine2.XDataEngine.Contains(Polyline, Key)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(Polyline As Polyline, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(Polyline, Key, Value)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(Polyline As Polyline, Transaction As Transaction, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(Transaction, Polyline, Key, Value)
        End Function

        ''' <summary>
        ''' Lê XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function GetXData(Polyline As Polyline, Key As String) As Object
            Return Engine2.XDataEngine.GetXData(Polyline, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(Polyline As Polyline, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(Polyline, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(Polyline As Polyline, Transaction As Transaction, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(Transaction, Polyline, Key)
        End Function

#End Region


#Region "Métodos XData em BlockReference"

        ''' <summary>
        ''' Verifica se existe um Xdata com chave específica
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function XDataContains(BlockReference As BlockReference, Key As String) As Boolean
            Return Engine2.XDataEngine.Contains(BlockReference, Key)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(BlockReference As BlockReference, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(BlockReference, Key, Value)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(BlockReference As BlockReference, Transaction As Transaction, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(Transaction, BlockReference, Key, Value)
        End Function

        ''' <summary>
        ''' Lê XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function GetXData(BlockReference As BlockReference, Key As String) As Object
            Return Engine2.XDataEngine.GetXData(BlockReference, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(BlockReference As BlockReference, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(BlockReference, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(BlockReference As BlockReference, Transaction As Transaction, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(Transaction, BlockReference, Key)
        End Function

#End Region

#Region "Métodos XData em DBText"

        ''' <summary>
        ''' Verifica se existe um Xdata com chave específica
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function XDataContains(DBText As DBText, Key As String) As Boolean
            Return Engine2.XDataEngine.Contains(DBText, Key)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(DBText As DBText, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(DBText, Key, Value)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(DBText As DBText, Transaction As Transaction, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(Transaction, DBText, Key, Value)
        End Function

        ''' <summary>
        ''' Lê XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function GetXData(DBText As DBText, Key As String) As Object
            Return Engine2.XDataEngine.GetXData(DBText, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(DBText As DBText, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(DBText, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(DBText As DBText, Transaction As Transaction, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(Transaction, DBText, Key)
        End Function

#End Region

#Region "Métodos XData em Wipeout"

        ''' <summary>
        ''' Verifica se existe um Xdata com chave específica
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function XDataContains(Wipeout As Wipeout, Key As String) As Boolean
            Return Engine2.XDataEngine.Contains(Wipeout, Key)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(Wipeout As Wipeout, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(Wipeout, Key, Value)
        End Function

        ''' <summary>
        ''' Grava XData na entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function SetXData(Wipeout As Wipeout, Transaction As Transaction, Key As String, Value As Object) As Boolean
            Return Engine2.XDataEngine.SetXData(Transaction, Wipeout, Key, Value)
        End Function

        ''' <summary>
        ''' Lê XData na entidade
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function GetXData(Wipeout As Wipeout, Key As String) As Object
            Return Engine2.XDataEngine.GetXData(Wipeout, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(Wipeout As Wipeout, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(Wipeout, Key)
        End Function

        ''' <summary>
        ''' Deleta XData
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DeleteXData(Wipeout As Wipeout, Transaction As Transaction, Key As String) As Boolean
            Return Engine2.XDataEngine.DeleteXdata(Transaction, Wipeout, Key)
        End Function

#End Region

#Region "Operação de valores em XML"

        ''' <summary>
        ''' Em uma string Xml verifica a existencia de chave e valor
        ''' </summary>
        ''' <param name="TagName">Nome da tag</param>
        ''' <param name="Value">Valor associado a tag</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function XmlContains(Xml As String, TagName As String, Value As String) As Boolean
            Return Engine2.Serialize.XmlContains(Xml, TagName, Value)
        End Function

        ''' <summary>
        ''' Em uma string Xml obtem um valor
        ''' </summary>
        ''' <param name="TagValue">Valor da tag</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function XmlGetValue(Xml As String, TagValue As String) As Object
            Return Engine2.Serialize.XmlGetValue(Xml, TagValue)
        End Function

        ''' <summary>
        ''' Atualiza o valor associado a uma tag no XML
        ''' </summary>
        ''' <param name="TagName">Nome da tag</param>
        ''' <param name="Value">Valor da tag</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function XmlUpdate(Xml As String, TagName As String, Value As Object) As Object
            Return Engine2.Serialize.XmlUpdate(Xml, TagName, Value)
        End Function

        ''' <summary>
        ''' Converte Xml em Classe
        ''' </summary>
        ''' <param name="ObjectItem">Classe</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function XmlToObject(Xml As String, ObjectItem As Object) As Object
            Return Engine2.Serialize.XmlToObject(Xml, ObjectItem)
        End Function

#End Region

#Region "Conversão de listas"

        ''' <summary>
        ''' Converte ObjectIdCollection em List(Of ObjectId)
        ''' </summary>
        ''' <param name="ObjectIdCollection"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ObjectIdCollectionToListOfObjectId(ObjectIdCollection As ObjectIdCollection) As List(Of ObjectId)
            Return Engine2.ConvertObject.ObjectIdCollectionToListOfObjectId(ObjectIdCollection)
        End Function

        ''' <summary>
        ''' Converte List(Of ObjectId) em ObjectIdCollection
        ''' </summary>
        ''' <param name="ListOfObjectId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ListOfObjectIdToObjectIdCollection(ListOfObjectId As List(Of ObjectId)) As ObjectIdCollection
            Return Engine2.ConvertObject.ListOfObjectIdToObjectIdCollection(ListOfObjectId)
        End Function

#End Region

    End Module

End Namespace