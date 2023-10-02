'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='
'LIMITE PARA ENTIDADES XDATA 8188 CARACTERES (Fonte: http://spiderinnet1.typepad.com/blog/2012/11/autocad-net-xdata-xdata-string-length-limit-physical-and-theoretical.html )
'=========================================================================================================='

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Namespace Engine2

    ''' <summary>
    '''  Classe para gravação de dados extendidos (XData)
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class XDataEngine

        ''' <summary>
        ''' Verifica se a Xdata existe na entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Contains(Entity As Entity, Key As String) As Boolean
            Try
                If IsNothing(Entity.GetXDataForApplication(Key)) = True Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Verifica se a Xdata existe na entidade
        ''' </summary>
        ''' <param name="DBObject">Database object</param>
        ''' <param name="Key">Chave</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Contains(DBObject As DBObject, Key As String) As Boolean
            Try
                If IsNothing(DBObject.GetXDataForApplication(Key)) = True Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Verifica se a Xdata existe na entidade
        ''' </summary>
        ''' <param name="ObjectId">ID da entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Contains(ObjectId As ObjectId, Key As String) As Boolean
            Try
                If IsNothing(Engine2.ConvertObject.ObjectIDToEntity(ObjectId).GetXDataForApplication(Key)) = True Then
                    Return False
                Else
                    Return True
                End If
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Obtem os dados estendidos
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Object</returns>
        ''' <remarks></remarks>
        Public Shared Function GetXData(Entity As Entity, Key As String) As Object
            Try
                Using ResultBuffer As ResultBuffer = Entity.GetXDataForApplication(Key)
                    If IsNothing(ResultBuffer) = False Then
                        For Each TypedValue As TypedValue In ResultBuffer
                            If TypedValue.TypeCode = 1000 Then
                                ResultBuffer.Dispose()
                                Return TypedValue.Value
                            End If
                        Next
                    End If
                    Return Nothing
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem os dados estendidos
        ''' </summary>
        ''' <param name="ObjectId">ID da entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Object</returns>
        ''' <remarks></remarks>
        Public Shared Function GetXData(ObjectId As ObjectId, Key As String) As Object
            Try
                Using ResultBuffer As ResultBuffer = ObjectId.ToEntity.GetXDataForApplication(Key)
                    If IsNothing(ResultBuffer) = False Then
                        For Each TypedValue As TypedValue In ResultBuffer
                            If TypedValue.TypeCode = 1000 Then
                                ResultBuffer.Dispose()
                                Return TypedValue.Value
                            End If
                        Next
                    End If
                    Return Nothing
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem os dados estendidos
        ''' </summary>
        ''' <param name="DBObject">Database object</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Object</returns>
        ''' <remarks></remarks>
        Public Shared Function GetXData(DBObject As DBObject, Key As String) As Object
            Using ResultBuffer As ResultBuffer = DBObject.GetXDataForApplication(Key)
                If IsNothing(ResultBuffer) = False Then
                    For Each TypedValue As TypedValue In ResultBuffer
                        If TypedValue.TypeCode = 1000 Then
                            ResultBuffer.Dispose()
                            Return TypedValue.Value
                        End If
                    Next
                End If
                Return Nothing
            End Using
        End Function

        ''' <summary>
        ''' Obtem a XData de um arquivo remoto
        ''' </summary>
        ''' <param name="FileName">Nome do arquivo</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="EntityType">Tipo de entidade</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetXData(FileName As String, Key As String, Optional EntityType As Object = Nothing) As Object
            If IsNothing(FileName) = False Then
                Try
                    Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                    Using DocumentLock As DocumentLock = Document.LockDocument()
                        Using Database As New Database(False, False)
                            With Database
                                .ReadDwgFile(FileName, FileOpenMode.OpenForReadAndReadShare, False, "")
                                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                                    Dim BlockTableRecord As BlockTableRecord = Transaction.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(Database), OpenMode.ForRead)
                                    For Each ObjectId As ObjectId In BlockTableRecord
                                        Dim Entity As Entity = Transaction.GetObject(ObjectId, OpenMode.ForRead)
                                        If Entity.GetType.Name = If(IsNothing(EntityType) = True, Entity.GetType.Name, EntityType.ToString) Then
                                            Using ResultBuffer As ResultBuffer = Entity.GetXDataForApplication(Key)
                                                If IsNothing(ResultBuffer) = False Then
                                                    For Each TypedValue As TypedValue In ResultBuffer
                                                        If TypedValue.TypeCode = 1000 Then
                                                            ResultBuffer.Dispose()
                                                            Return TypedValue.Value
                                                        End If
                                                    Next
                                                End If
                                            End Using
                                        End If
                                    Next
                                    Transaction.Dispose()
                                End Using
                                .Dispose()
                            End With
                        End Using
                    End Using
                    Return Nothing
                Catch
                    Return Nothing
                End Try
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Grava os dados estendidos 
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function SetXData(Entity As Entity, Key As String, Value As String) As Boolean
            Try
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction
                        Try
                            Dim DBObject As DBObject = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite)
                            Dim RegAppTable As RegAppTable = Transaction.GetObject(HostApplicationServices.WorkingDatabase.RegAppTableId, OpenMode.ForWrite)
                            If RegAppTable.Has(Key) = False Then
                                Using RegAppTableRecord As New RegAppTableRecord
                                    RegAppTable.UpgradeOpen()
                                    RegAppTableRecord.Name = Key
                                    RegAppTable.Add(RegAppTableRecord)
                                    Transaction.AddNewlyCreatedDBObject(RegAppTableRecord, True)
                                End Using
                            End If
                            Using ResultBuffer As New ResultBuffer(New TypedValue(1001, Key), New TypedValue(1000, Value))
                                DBObject.XData = ResultBuffer
                            End Using
                            Transaction.Commit()
                            Return True
                        Catch
                            Transaction.Abort()
                            Return False
                        End Try
                    End Using
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Grava os dados estendidos 
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function SetXData(Transaction As Transaction, Entity As Entity, Key As String, Value As String) As Boolean
            Try
                Dim DBObject As DBObject = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite)
                Dim RegAppTable As RegAppTable = Transaction.GetObject(HostApplicationServices.WorkingDatabase.RegAppTableId, OpenMode.ForWrite)
                If RegAppTable.Has(Key) = False Then
                    Using RegAppTableRecord As New RegAppTableRecord
                        RegAppTable.UpgradeOpen()
                        RegAppTableRecord.Name = Key
                        RegAppTable.Add(RegAppTableRecord)
                        Transaction.AddNewlyCreatedDBObject(RegAppTableRecord, True)
                    End Using
                End If
                Using ResultBuffer As New ResultBuffer(New TypedValue(1001, Key), New TypedValue(1000, Value))
                    DBObject.XData = ResultBuffer
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Grava os dados estendidos 
        ''' </summary>
        ''' <param name="ObjectId">ID da entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function SetXData(ObjectId As ObjectId, Key As String, Value As String) As Boolean
            Try
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction
                        Try
                            Dim DBObject As DBObject = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                            Dim RegAppTable As RegAppTable = Transaction.GetObject(HostApplicationServices.WorkingDatabase.RegAppTableId, OpenMode.ForWrite)
                            If RegAppTable.Has(Key) = False Then
                                Using RegAppTableRecord As New RegAppTableRecord
                                    RegAppTable.UpgradeOpen()
                                    RegAppTableRecord.Name = Key
                                    RegAppTable.Add(RegAppTableRecord)
                                    Transaction.AddNewlyCreatedDBObject(RegAppTableRecord, True)
                                End Using
                            End If
                            Using ResultBuffer As New ResultBuffer(New TypedValue(1001, Key), New TypedValue(1000, Value))
                                DBObject.XData = ResultBuffer
                            End Using
                            Transaction.Commit()
                            Return True
                        Catch
                            Transaction.Abort()
                            Return False
                        End Try
                    End Using
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Grava os dados estendidos 
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ObjectId">ID da entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function SetXData(Transaction As Transaction, ObjectId As ObjectId, Key As String, Value As String) As Boolean
            Try
                Dim DBObject As DBObject = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                Dim RegAppTable As RegAppTable = Transaction.GetObject(HostApplicationServices.WorkingDatabase.RegAppTableId, OpenMode.ForWrite)
                If RegAppTable.Has(Key) = False Then
                    Using RegAppTableRecord As New RegAppTableRecord
                        RegAppTable.UpgradeOpen()
                        RegAppTableRecord.Name = Key
                        RegAppTable.Add(RegAppTableRecord)
                        Transaction.AddNewlyCreatedDBObject(RegAppTableRecord, True)
                    End Using
                End If
                Using ResultBuffer As New ResultBuffer(New TypedValue(1001, Key), New TypedValue(1000, Value))
                    DBObject.XData = ResultBuffer
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Grava os dados estendidos 
        ''' </summary>
        ''' <param name="DBObject">Database object</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function SetXData(DBObject As DBObject, Key As String, Value As String) As Boolean
            Try
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction
                        Try
                            DBObject = Transaction.GetObject(DBObject.ObjectId, OpenMode.ForWrite)
                            Dim RegAppTable As RegAppTable = Transaction.GetObject(HostApplicationServices.WorkingDatabase.RegAppTableId, OpenMode.ForWrite)
                            If RegAppTable.Has(Key) = False Then
                                Using RegAppTableRecord As New RegAppTableRecord
                                    RegAppTable.UpgradeOpen()
                                    RegAppTableRecord.Name = Key
                                    RegAppTable.Add(RegAppTableRecord)
                                    Transaction.AddNewlyCreatedDBObject(RegAppTableRecord, True)
                                End Using
                            End If
                            Using ResultBuffer As New ResultBuffer(New TypedValue(1001, Key), New TypedValue(1000, Value))
                                DBObject.XData = ResultBuffer
                            End Using
                            Transaction.Commit()
                            Return True
                        Catch
                            Transaction.Abort()
                            Return False
                        End Try
                    End Using
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Grava os dados estendidos 
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DBObject">Database object</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function SetXData(Transaction As Transaction, DBObject As DBObject, Key As String, Value As String) As Boolean
            Try
                DBObject = Transaction.GetObject(DBObject.ObjectId, OpenMode.ForWrite)
                Dim RegAppTable As RegAppTable = Transaction.GetObject(HostApplicationServices.WorkingDatabase.RegAppTableId, OpenMode.ForWrite)
                If RegAppTable.Has(Key) = False Then
                    Using RegAppTableRecord As New RegAppTableRecord
                        RegAppTable.UpgradeOpen()
                        RegAppTableRecord.Name = Key
                        RegAppTable.Add(RegAppTableRecord)
                        Transaction.AddNewlyCreatedDBObject(RegAppTableRecord, True)
                    End Using
                End If
                Using ResultBuffer As New ResultBuffer(New TypedValue(1001, Key), New TypedValue(1000, Value))
                    DBObject.XData = ResultBuffer
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Deleta dados estendidos (XData)
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteXdata(Entity As Entity, Key As String) As Boolean
            Try
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction
                        Try
                            Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite)
                            Entity.XData = New ResultBuffer(New TypedValue(1001, Key))
                            Transaction.Commit()
                            Return True
                        Catch
                            Transaction.Abort()
                            Return False
                        End Try
                    End Using
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Deleta dados estendidos (XData)
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteXdata(Transaction As Transaction, Entity As Entity, Key As String) As Boolean
            Try
                Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite)
                Entity.XData = New ResultBuffer(New TypedValue(1001, Key))
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Deleta dados estendidos (XData)
        ''' </summary>
        ''' <param name="ObjectId">ID da entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteXdata(ObjectId As ObjectId, Key As String) As Boolean
            Try
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction
                        Try
                            Dim Entity As Entity = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                            Entity.XData = New ResultBuffer(New TypedValue(1001, Key))
                            Transaction.Commit()
                            Return True
                        Catch
                            Transaction.Abort()
                            Return False
                        End Try
                    End Using
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Deleta dados estendidos (XData)
        ''' </summary>
        ''' <param name="Transaction">Transaction</param>
        ''' <param name="ObjectId">ID da entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteXdata(Transaction As Transaction, ObjectId As ObjectId, Key As String) As Boolean
            Try
                Dim Entity As Entity = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                Entity.XData = New ResultBuffer(New TypedValue(1001, Key))
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Deleta dados estendidos (XData)
        ''' </summary>
        ''' <param name="DBObject">Database object</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteXdata(DBObject As DBObject, Key As String) As Boolean
            Try
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction
                        Try
                            DBObject = Transaction.GetObject(DBObject.ObjectId, OpenMode.ForWrite)
                            DBObject.XData = New ResultBuffer(New TypedValue(1001, Key))
                            Transaction.Commit()
                            Return True
                        Catch
                            Transaction.Abort()
                            Return False
                        End Try
                    End Using
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Deleta dados estendidos (XData)
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DBObject">Database object</param>
        ''' <param name="Key">Chave</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteXdata(Transaction As Transaction, DBObject As DBObject, Key As String) As Boolean
            Try
                DBObject = Transaction.GetObject(DBObject.ObjectId, OpenMode.ForWrite)
                DBObject.XData = New ResultBuffer(New TypedValue(1001, Key))
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Exclui todos os dados extendidos de uma entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteAllXdata(Entity As Entity) As Boolean
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim appNames As New ArrayList
            Dim Applications As ArrayList
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    Try
                        Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
                        If IsNothing(Entity.XData) = False Then
                            Entity.UpgradeOpen()
                            Applications = GetApplications(Entity)
                            For Each Application As String In Applications
                                Entity.XData = New ResultBuffer(New TypedValue(1001, Application))
                            Next
                        End If
                        Transaction.Commit()
                        Return True
                    Catch
                        Transaction.Abort()
                        Return False
                    End Try
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Exclui todos os dados extendidos de uma entidade
        ''' </summary>
        ''' <param name="DBObject">DBObject</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteAllXdata(DBObject As DBObject) As Boolean
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim appNames As New ArrayList
            Dim Applications As ArrayList
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    Try
                        DBObject = Transaction.GetObject(DBObject.ObjectId, OpenMode.ForRead)
                        If IsNothing(DBObject.XData) = False Then
                            DBObject.UpgradeOpen()
                            Applications = GetApplications(DBObject)
                            For Each Application As String In Applications
                                DBObject.XData = New ResultBuffer(New TypedValue(1001, Application))
                            Next
                        End If
                        Transaction.Commit()
                        Return True
                    Catch
                        Transaction.Abort()
                        Return False
                    End Try
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Exclui todos os dados extendidos de uma entidade
        ''' </summary>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteAllXdata(ObjectId As ObjectId) As Boolean
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim appNames As New ArrayList
            Dim Applications As ArrayList
            Dim Entity As Entity
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    Try
                        Entity = Transaction.GetObject(ObjectId, OpenMode.ForRead)
                        If IsNothing(Entity.XData) = False Then
                            Entity.UpgradeOpen()
                            Applications = GetApplications(ObjectId)
                            For Each Application As String In Applications
                                Entity.XData = New ResultBuffer(New TypedValue(1001, Application))
                            Next
                        End If
                        Transaction.Commit()
                        Return True
                    Catch
                        Transaction.Abort()
                        Return False
                    End Try
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Exclui todos os dados extendidos de uma entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Entity">Entidade</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteAllXdata(Transaction As Transaction, Entity As Entity) As Boolean
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim appNames As New ArrayList
            Dim Applications As ArrayList
            Try
                Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
                If IsNothing(Entity.XData) = False Then
                    Entity.UpgradeOpen()
                    Applications = GetApplications(Entity)
                    For Each Application As String In Applications
                        Entity.XData = New ResultBuffer(New TypedValue(1001, Application))
                    Next
                End If
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Exclui todos os dados extendidos de uma entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DBObject">DBObject</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteAllXdata(Transaction As Transaction, DBObject As DBObject) As Boolean
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim appNames As New ArrayList
            Dim Applications As ArrayList
            Try
                DBObject = Transaction.GetObject(DBObject.ObjectId, OpenMode.ForRead)
                If IsNothing(DBObject.XData) = False Then
                    DBObject.UpgradeOpen()
                    Applications = GetApplications(DBObject)
                    For Each Application As String In Applications
                        DBObject.XData = New ResultBuffer(New TypedValue(1001, Application))
                    Next
                End If
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Exclui todos os dados extendidos de uma entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteAllXdata(Transaction As Transaction, ObjectId As ObjectId) As Boolean
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim appNames As New ArrayList
            Dim Applications As ArrayList
            Dim Entity As Entity = ObjectId.ToEntity
            Try
                Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
                If IsNothing(Entity.XData) = False Then
                    Entity.UpgradeOpen()
                    Applications = GetApplications(ObjectId)
                    For Each Application As String In Applications
                        Entity.XData = New ResultBuffer(New TypedValue(1001, Application))
                    Next
                End If
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Retorna todas as aplicações Xdata de uma entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetApplications(Entity As Entity) As ArrayList
            GetApplications = New ArrayList
            If IsNothing(Entity.XData) = False Then
                For Each TypedValue As TypedValue In Entity.XData.AsArray
                    If TypedValue.TypeCode = 1001 Then
                        GetApplications.Add(TypedValue.Value.ToString)
                    End If
                Next
                GetApplications.Sort()
            End If
            Return GetApplications
        End Function

        ''' <summary>
        ''' Retorna todas as aplicações Xdata de uma entidade
        ''' </summary>
        ''' <param name="DBObject">DBObject</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetApplications(DBObject As DBObject) As ArrayList
            GetApplications = New ArrayList
            If IsNothing(DBObject.XData) = False Then
                For Each TypedValue As TypedValue In DBObject.XData.AsArray
                    If TypedValue.TypeCode = 1001 Then
                        GetApplications.Add(TypedValue.Value.ToString)
                    End If
                Next
                GetApplications.Sort()
            End If
            Return GetApplications
        End Function

        ''' <summary>
        ''' Retorna todas as aplicações Xdata de uma entidade
        ''' </summary>
        ''' <param name="ObjectId">ObjectId</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetApplications(ObjectId As ObjectId) As ArrayList
            GetApplications = New ArrayList
            If IsNothing(ObjectId.ToEntity.XData) = False Then
                For Each TypedValue As TypedValue In ObjectId.ToEntity.XData.AsArray
                    If TypedValue.TypeCode = 1001 Then
                        GetApplications.Add(TypedValue.Value.ToString)
                    End If
                Next
                GetApplications.Sort()
            End If
            Return GetApplications
        End Function

    End Class
End Namespace
