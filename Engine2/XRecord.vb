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
Imports Autodesk.AutoCAD.DatabaseServices.OpenMode

Namespace Engine2

    ''' <summary>
    ''' Classe para gravação de dados extendidos (XRecord)
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class XRecordEngine

        ''' <summary>
        ''' Grava dados estendidos (XRecord)
        ''' </summary>
        ''' <param name="Document">Documento do AutoCAD</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor a ser gravado</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function SetXRecord(Document As Document, Key As String, Value As String, Optional Transaction As Transaction = Nothing) As Boolean
            SetXRecord = False
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim DBDictionary As DBDictionary
            Dim Xrecord As New Xrecord
            If IsNothing(Transaction) = False Then
                DBDictionary = Transaction.GetObject(Document.Database.NamedObjectsDictionaryId, OpenMode.ForWrite, False)
                Xrecord = New Xrecord()
                Xrecord.Data = New ResultBuffer(New TypedValue(CInt(DxfCode.Text), Value))
                DBDictionary.SetAt(Key, Xrecord)
                Transaction.AddNewlyCreatedDBObject(Xrecord, True)
                SetXRecord = True
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Transaction = Document.Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            SetXRecord = SetXRecord(Document, Key, Value, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                        End Try
                    End Using
                End Using
            End If
            Return SetXRecord
        End Function

        ''' <summary>
        ''' Obtem dados estendidos (XRecord)
        ''' </summary>
        ''' <param name="Document">Documento do AutoCAD</param>
        ''' <param name="key">Chave</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>Object</returns>
        ''' <remarks></remarks>
        Public Shared Function GetXRecord(Document As Document, key As String, Optional Transaction As Transaction = Nothing) As Object
            GetXRecord = Nothing
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim ObjectId As ObjectId
            Dim Xrecord As Xrecord
            Dim DBDictionary As DBDictionary
            If IsNothing(Transaction) = False Then
                DBDictionary = Transaction.GetObject(Document.Database.NamedObjectsDictionaryId, OpenMode.ForRead)
                ObjectId = DBDictionary.GetAt(key)
                Xrecord = DirectCast(Transaction.GetObject(ObjectId, OpenMode.ForRead), Xrecord)
                For Each TypedValue As TypedValue In Xrecord.Data
                    If TypedValue.TypeCode = DxfCode.Text Then
                        GetXRecord = TypedValue.Value
                    End If
                Next
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Transaction = Document.Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            GetXRecord = GetXRecord(Document, key, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                        End Try
                    End Using
                End Using
            End If
            Return GetXRecord
        End Function

        ''' <summary>
        ''' Grava dados estendidos (XRecord)
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Value">Valor a ser gravado</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function SetXRecord(Entity As Entity, Key As String, Value As String, Optional Transaction As Transaction = Nothing) As Boolean
            SetXRecord = False
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim DBDictionary As DBDictionary
            Dim Xrecord As Xrecord
            Dim EntityID As ObjectId
            Dim ExtDict As DBDictionary
            If IsNothing(Transaction) = False Then
                Xrecord = Nothing
                EntityID = Entity.ExtensionDictionary
                If EntityID = ObjectId.Null Then
                    Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForWrite)
                    Entity.CreateExtensionDictionary()
                    EntityID = Entity.ExtensionDictionary
                End If
                ExtDict = Transaction.GetObject(EntityID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, False)
                If ExtDict.Contains(Key) Then
                    DBDictionary = Transaction.GetObject(ExtDict.GetAt(Key), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, False)
                Else
                    DBDictionary = New DBDictionary
                    ExtDict.SetAt(Key, DBDictionary)
                    Transaction.AddNewlyCreatedDBObject(DBDictionary, True)
                End If
                If DBDictionary.Contains(Key) Then
                    Xrecord = Transaction.GetObject(DBDictionary.GetAt(Key), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, False)
                Else
                    Xrecord = New Xrecord
                    DBDictionary.SetAt(Key, Xrecord)
                    Transaction.AddNewlyCreatedDBObject(Xrecord, True)
                End If
                Xrecord.Data = New ResultBuffer(New TypedValue(CInt(DxfCode.Text), Value))
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Transaction = Document.Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            SetXRecord = SetXRecord(Document, Key, Value, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                        End Try
                    End Using
                End Using
            End If
            Return SetXRecord
        End Function

        ''' <summary>
        ''' Obtem dados estendidos (XRecord)
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Key">Chave</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>Object</returns>
        ''' <remarks></remarks>
        Public Shared Function GetXRecord(Entity As Entity, Key As String, Optional Transaction As Transaction = Nothing) As Object
            GetXRecord = Nothing
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim Xrecord As Xrecord
            Dim ObjectId As ObjectId
            Dim ExtDict As DBDictionary
            Dim DictId As ObjectId
            Dim MyDict As DBDictionary
            If IsNothing(Transaction) = False Then
                Xrecord = Nothing
                ObjectId = Entity.ExtensionDictionary
                If ObjectId.IsValid = True Then
                    ExtDict = Transaction.GetObject(ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, False)
                    If ExtDict.Contains(Key) Then
                        DictId = ExtDict.GetAt(Key)
                        MyDict = Transaction.GetObject(DictId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, False)
                        Xrecord = Transaction.GetObject(MyDict.GetAt(Key), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, False)
                        If IsNothing(Xrecord) = False Then
                            If IsNothing(Xrecord.Data) = False Then
                                For Each TypedValue As TypedValue In Xrecord.Data
                                    If TypedValue.TypeCode = DxfCode.Text Then
                                        GetXRecord = TypedValue.Value
                                        Exit For
                                    End If
                                Next
                            Else
                                GetXRecord = Nothing
                            End If
                        Else
                            GetXRecord = Nothing
                        End If
                    Else
                        GetXRecord = Nothing
                    End If
                Else
                    GetXRecord = Nothing
                End If
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Transaction = Document.Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            GetXRecord = GetXRecord(Entity, Key, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                        End Try
                    End Using
                End Using
            End If
            Return GetXRecord
        End Function

        ''' <summary>
        ''' Deleta dados estendidos (XRecord)
        ''' </summary>
        ''' <param name="Document">Documento do AutoCAD</param>
        ''' <param name="key">Chave</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteXRecord(Document As Document, key As String, Optional Transaction As Transaction = Nothing) As Boolean
            DeleteXRecord = False
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim ObjectId As ObjectId
            Dim DBDictionary As DBDictionary
            If IsNothing(Transaction) = False Then
                DBDictionary = DirectCast(Transaction.GetObject(Database.NamedObjectsDictionaryId, OpenMode.ForWrite), DBDictionary)
                If DBDictionary.Contains(key) = True Then
                    ObjectId = DBDictionary.GetAt(key)
                    DBDictionary.Remove(ObjectId)
                    DBDictionary.Dispose()
                    DeleteXRecord = True
                End If
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Transaction = Document.Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            DeleteXRecord = DeleteXRecord(Document, key, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                        End Try
                    End Using
                End Using
            End If
            Return DeleteXRecord
        End Function

        ''' <summary>
        ''' Deleta dados estendidos (XRecord)
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="key">Chave</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteXRecord(Entity As Entity, key As String, Optional Transaction As Transaction = Nothing) As Boolean
            DeleteXRecord = False
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim ObjectId As ObjectId
            If IsNothing(Transaction) = False Then
                Entity = TryCast(Transaction.GetObject(Entity.Id, OpenMode.ForWrite), Entity)
                If Entity.ExtensionDictionary <> ObjectId.Null Then
                    Using DBDictionary As DBDictionary = Transaction.GetObject(Entity.ExtensionDictionary, OpenMode.ForWrite)
                        If DBDictionary.Contains(key) = True Then
                            ObjectId = DBDictionary.GetAt(key)
                            DBDictionary.Remove(ObjectId)
                            DeleteXRecord = True
                        Else
                            DeleteXRecord = True
                        End If
                    End Using
                Else
                    DeleteXRecord = True
                End If
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Transaction = Document.Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            DeleteXRecord = DeleteXRecord(Entity, key, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                        End Try
                    End Using
                End Using
            End If
            Return DeleteXRecord
        End Function

    End Class

End Namespace
