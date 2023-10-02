Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.IO
Imports Autodesk.AutoCAD.EditorInput
Imports System.Drawing
Imports Autodesk.AutoCAD.Geometry
Imports System.Windows.Media
Imports Autodesk.AutoCAD.Windows.Data
Imports System.Windows.Media.Imaging
Imports System.Collections
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Internal

'14/03/19 MODIFICADO E NÃO REPASSADO AS OUTRAS VERSÕES POR CONTA DE TESTE

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2019
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Classe para inserção de blocos
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Block

        ''' <summary>
        ''' Filtro de blocos
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eBlockFilter
            XRef = 0
            Block = 1
            All = 3
        End Enum

        ''' <summary>
        ''' Versões
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eFileVersions
            Acad_2018_2019_2020 = 9
            Acad_2013_2014_2015_2016_2017 = 8
            Acad_2010_2011_2012 = 7
            Acad_2007_2008_2009 = 6
            Acad_2004_2005_2006 = 5
            Acad_2000_200i_2002 = 4
            Acad_Release_14 = 3
            Acad_Release_13 = 2
            Acad_Release_12_11 = 1
            Acad_Release_9 = 0
        End Enum

        ''' <summary>
        ''' Verifica se um bloco existe
        ''' </summary>
        ''' <param name="BlockName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BlockExists(BlockName As String) As Boolean
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim BlockTable As BlockTable
            Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                Return BlockTable.Has(BlockName)
            End Using
        End Function

        ''' <summary>
        ''' Verifica se um bloco existe
        ''' </summary>
        ''' <param name="Transaction"></param>
        ''' <param name="BlockName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BlockExists(Transaction As Transaction, BlockName As String) As Boolean
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim BlockTable As BlockTable
            BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
            Return BlockTable.Has(BlockName)
        End Function

        ''' <summary>
        ''' Verifica se um arquivo é compatível 
        ''' </summary>
        ''' <param name="FileName"></param>
        ''' <param name="FileVersion"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CheckFileVersion(FileName As String, FileVersion As eFileVersions) As Boolean
            'AC1032 AutoCAD 2018/2019/2020
            'AC1027 AutoCAD 2013/2014/2015/2016/2017
            'AC1024 AutoCAD 2010/2011/2012 
            'AC1021 AutoCAD 2007/2008/2009 
            'AC1018 AutoCAD 2004/2005/2006 
            'AC1015 AutoCAD 2000/2000i/2002 
            'AC1014 Release 14 
            'AC1012 Release 13 
            'AC1009 Release 11/12 
            'AC1006 Release 10 
            'AC1004 Release 9 
            'AC1003 Version 2.60 
            'AC1002 Version 2.50 
            'AC1001 Version 2.22 
            'AC2.22 Version 2.22 
            'AC2.21 Version 2.21 
            'AC2.10 Version 2.10 
            'AC1.50 Version 2.05 
            'AC1.40 Version 1.40 
            'AC1.2 Version 1.2 
            'MC0.0 Version 1.0
            Dim CurrentVersion As String = ""
            Using reader As New StreamReader(FileName)
                CurrentVersion = reader.ReadLine().Substring(0, 6)
                reader.Dispose()
            End Using
            Select Case FileVersion
                Case eFileVersions.Acad_Release_9
                    If CurrentVersion = "AC1004" Then
                        Return True
                    End If
                Case eFileVersions.Acad_Release_12_11
                    If CurrentVersion = "AC1009" Or CurrentVersion = "AC1004" Then
                        Return True
                    End If
                Case eFileVersions.Acad_Release_13
                    If CurrentVersion = "AC1012" Or CurrentVersion = "AC1009" Or CurrentVersion = "AC1004" Then
                        Return True
                    End If
                Case eFileVersions.Acad_Release_14
                    If CurrentVersion = "AC1014" Or CurrentVersion = "AC1012" Or CurrentVersion = "AC1009" Or CurrentVersion = "AC1004" Then
                        Return True
                    End If
                Case eFileVersions.Acad_2000_200i_2002
                    If CurrentVersion = "AC1015" Or CurrentVersion = "AC1014" Or CurrentVersion = "AC1012" Or CurrentVersion = "AC1009" Or CurrentVersion = "AC1004" Then
                        Return True
                    End If
                Case eFileVersions.Acad_2004_2005_2006
                    If CurrentVersion = "AC1018" Or CurrentVersion = "AC1015" Or CurrentVersion = "AC1014" Or CurrentVersion = "AC1012" Or CurrentVersion = "AC1009" Or CurrentVersion = "AC1004" Then
                        Return True
                    End If
                Case eFileVersions.Acad_2007_2008_2009
                    If CurrentVersion = "AC1021" Or CurrentVersion = "AC1018" Or CurrentVersion = "AC1015" Or CurrentVersion = "AC1014" Or CurrentVersion = "AC1012" Or CurrentVersion = "AC1009" Or CurrentVersion = "AC1004" Then
                        Return True
                    End If
                Case eFileVersions.Acad_2010_2011_2012
                    If CurrentVersion = "AC1024" Or CurrentVersion = "AC1021" Or CurrentVersion = "AC1018" Or CurrentVersion = "AC1015" Or CurrentVersion = "AC1014" Or CurrentVersion = "AC1012" Or CurrentVersion = "AC1009" Or CurrentVersion = "AC1004" Then
                        Return True
                    End If
                Case eFileVersions.Acad_2013_2014_2015_2016_2017
                    If CurrentVersion = "AC1027" Or CurrentVersion = "AC1024" Or CurrentVersion = "AC1021" Or CurrentVersion = "AC1018" Or CurrentVersion = "AC1015" Or CurrentVersion = "AC1014" Or CurrentVersion = "AC1012" Or CurrentVersion = "AC1009" Or CurrentVersion = "AC1004" Then
                        Return True
                    End If
                Case eFileVersions.Acad_2018_2019_2020
                    If CurrentVersion = "AC1032" Or CurrentVersion = "AC1027" Or CurrentVersion = "AC1024" Or CurrentVersion = "AC1021" Or CurrentVersion = "AC1018" Or CurrentVersion = "AC1015" Or CurrentVersion = "AC1014" Or CurrentVersion = "AC1012" Or CurrentVersion = "AC1009" Or CurrentVersion = "AC1004" Then
                        Return True
                    End If
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' Insere o bloco interno
        ''' </summary>
        ''' <param name="Block">Bloco</param>
        ''' <param name="Position">Posição</param> 
        ''' <param name="CloneAttributesValue">Determina se o valor do atributo será transferido para o novo bloco</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>BlockReference</returns>
        ''' <remarks></remarks>
        Public Shared Function Insert(Block As BlockReference, Optional Position As Object = Nothing, Optional CloneAttributesValue As Boolean = False, Optional Transaction As Transaction = Nothing) As BlockReference
            Insert = Nothing
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim BlockTable As BlockTable
            Dim BlockID As ObjectId
            Dim Workspace As BlockTableRecord
            Dim AttributeDefinition As AttributeDefinition
            Dim DBObject As DBObject
            Dim BlockTableRecord As BlockTableRecord
            Dim AttributeReference As AttributeReference
            Dim AnnotativeStates As AnnotativeStates
            Dim ObjectContextManager As ObjectContextManager
            Dim ObjectContextCollection As ObjectContextCollection
            Try
                If IsNothing(Transaction) = False Then
                    If IsNothing(Position) = True Then
                        Position = Block.Position
                    End If
                    BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                    BlockID = BlockTable(Block.Name)
                    Workspace = Transaction.GetObject(Document.Database.CurrentSpaceId, OpenMode.ForWrite)
                    Insert = New BlockReference(Position, BlockID)
                    Insert.Layer = Block.Layer
                    Insert.Rotation = Block.Rotation
                    Insert.ScaleFactors = Block.ScaleFactors
                    Workspace.AppendEntity(Insert)
                    Insert.SetDatabaseDefaults()
                    AnnotativeStates = Insert.Annotative
                    If AnnotativeStates = Autodesk.AutoCAD.DatabaseServices.AnnotativeStates.True Then
                        ObjectContextManager = Database.ObjectContextManager
                        ObjectContextCollection = ObjectContextManager.GetContextCollection("ACDB_ANNOTATIONSCALES")
                        ObjectContexts.AddContext(Insert, ObjectContextCollection.CurrentContext)
                    End If
                    Transaction.AddNewlyCreatedDBObject(Insert, True)
                    BlockTableRecord = Transaction.GetObject(Insert.BlockTableRecord, OpenMode.ForRead)
                    If BlockTableRecord.HasAttributeDefinitions = True Then
                        For Each objID As ObjectId In BlockTableRecord
                            DBObject = Transaction.GetObject(objID, OpenMode.ForRead)
                            If DBObject.GetType.Name = "AttributeDefinition" Then
                                AttributeDefinition = DBObject
                                If AttributeDefinition.Constant = False Then
                                    AttributeReference = New AttributeReference
                                    Using AttributeReference
                                        AttributeReference.SetAttributeFromBlock(AttributeDefinition, Insert.BlockTransform)
                                        AttributeReference.Position = AttributeDefinition.Position.TransformBy(Insert.BlockTransform)
                                        If CloneAttributesValue = True Then
                                            AttributeReference.TextString = ReadAttribute(Block, AttributeDefinition.Tag, Transaction)
                                        Else
                                            AttributeReference.TextString = AttributeDefinition.TextString
                                        End If
                                        AttributeReference.Tag = AttributeDefinition.Tag
                                        Insert.AttributeCollection.AppendAttribute(AttributeReference)
                                        Transaction.AddNewlyCreatedDBObject(AttributeReference, True)
                                    End Using
                                End If
                            End If
                        Next
                    End If
                Else
                    Using DocumentLock As DocumentLock = Document.LockDocument
                        Transaction = Document.TransactionManager.StartTransaction
                        Using Transaction
                            Try
                                Insert = Insert(Block, Position, CloneAttributesValue, Transaction)
                                Transaction.Commit()
                            Catch
                                Transaction.Abort()
                                Exit Try
                            End Try
                        End Using
                    End Using
                End If
                Return Insert
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Insere o bloco interno
        ''' </summary>
        ''' <param name="BlockName">Nome do bloco</param>
        ''' <param name="Position">Posição</param>
        ''' <param name="ScaleFactors">Fator de escala</param>
        ''' <param name="Rotation">Rotação</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <param name="AngleFormat">Formato informado para a rotação</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Insert(BlockName As String, Position As Point3d, ScaleFactors As Scale3d, Optional Rotation As Double = 0, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = False, Optional AngleFormat As Engine2.Geometry.eAngleFormat = Geometry.eAngleFormat.Degrees, Optional Transaction As Transaction = Nothing) As BlockReference
            Insert = Nothing
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim BlockTable As BlockTable
            Dim BlockID As ObjectId
            Dim Workspace As BlockTableRecord
            Dim AttributeDefinition As AttributeDefinition
            Dim DBObject As DBObject
            Dim BlockTableRecord As BlockTableRecord
            Dim AttributeReference As AttributeReference
            Dim AnnotativeStates As AnnotativeStates
            Dim ObjectContextManager As ObjectContextManager
            Dim ObjectContextCollection As ObjectContextCollection
            Try
                If IsNothing(Transaction) = False Then
                    BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                    BlockID = BlockTable(BlockName)
                    Workspace = Transaction.GetObject(Document.Database.CurrentSpaceId, OpenMode.ForWrite)
                    Insert = New BlockReference(Position, BlockID)
                    If IsNothing(Layer) = False Then
                        Insert.Layer = Layer
                    End If
                    If AngleFormat = Geometry.eAngleFormat.Degrees Then
                        Insert.Rotation = Engine2.Geometry.DegreeToRadian(Rotation)
                    Else
                        Insert.Rotation = Rotation
                    End If
                    Insert.ScaleFactors = ScaleFactors
                    Workspace.AppendEntity(Insert)
                    Insert.SetDatabaseDefaults()
                    AnnotativeStates = Insert.Annotative
                    If AnnotativeStates = Autodesk.AutoCAD.DatabaseServices.AnnotativeStates.True Then
                        ObjectContextManager = Database.ObjectContextManager
                        ObjectContextCollection = ObjectContextManager.GetContextCollection("ACDB_ANNOTATIONSCALES")
                        ObjectContexts.AddContext(Insert, ObjectContextCollection.CurrentContext)
                    End If
                    Transaction.AddNewlyCreatedDBObject(Insert, True)
                    BlockTableRecord = Transaction.GetObject(Insert.BlockTableRecord, OpenMode.ForRead)
                    If BlockTableRecord.HasAttributeDefinitions = True Then
                        For Each objID As ObjectId In BlockTableRecord
                            DBObject = Transaction.GetObject(objID, OpenMode.ForRead)
                            If DBObject.GetType.Name = "AttributeDefinition" Then
                                AttributeDefinition = DBObject
                                If AttributeDefinition.Constant = False Then
                                    AttributeReference = New AttributeReference
                                    Using AttributeReference
                                        AttributeReference.SetAttributeFromBlock(AttributeDefinition, Insert.BlockTransform)
                                        AttributeReference.Position = AttributeDefinition.Position.TransformBy(Insert.BlockTransform)
                                        AttributeReference.TextString = AttributeDefinition.TextString
                                        AttributeReference.Tag = AttributeDefinition.Tag
                                        Insert.AttributeCollection.AppendAttribute(AttributeReference)
                                        Transaction.AddNewlyCreatedDBObject(AttributeReference, True)
                                    End Using
                                End If
                            End If
                        Next
                    End If
                Else
                    Using DocumentLock As DocumentLock = Document.LockDocument
                        Transaction = Document.TransactionManager.StartTransaction
                        Using Transaction
                            Try
                                Insert = Insert(BlockName, Position, ScaleFactors, Rotation, Layer, UseUCS, AngleFormat, Transaction)
                                Transaction.Commit()
                            Catch
                                Transaction.Abort()
                                Exit Try
                            End Try
                        End Using
                    End Using
                End If
                Return Insert
            Catch
                Return Nothing
            End Try
        End Function


        ''' <summary>
        ''' Insert 
        ''' </summary>
        ''' <param name="Path">Nome e caminho do arquivo a ser inserido</param>
        ''' <param name="Position">Ponto de inserção</param>
        ''' <param name="ScaleFactors">Escala</param>
        ''' <param name="Rotation">Rotação</param>
        ''' <param name="BlockName">Nome do bloco</param>
        ''' <param name="Document">Documento</param>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ActiveDocument">Determina se o documento será ativado</param>
        ''' <param name="LayoutName">Nome do layout onde será inserido o bloco</param>
        ''' <param name="Layer">Camada</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Insert(Path As FileInfo, Position As Point3d, ScaleFactors As Scale3d, Optional Rotation As Double = 0, Optional BlockName As String = "", Optional Layer As Object = Nothing, Optional LayoutName As Object = Nothing, Optional Document As Document = Nothing, Optional Transaction As Transaction = Nothing, Optional ActiveDocument As Boolean = True) As BlockReference
            Try
                Dim Editor As Editor
                Dim Database As Database
                Dim ObjectId As ObjectId
                Dim BlockTableRecord As BlockTableRecord = Nothing
                Dim BlockReference As BlockReference = Nothing
                Dim LayoutManager As LayoutManager
                Dim LayoutId As ObjectId
                Dim Layout As Layout
                Dim BlockTable As BlockTable
                Dim TypedValueCollection As New List(Of TypedValue)
                Dim AttributeDefinitionCollection As New List(Of AttributeDefinition)
                Dim AttributeDefinition As AttributeDefinition
                Dim AttributeReference As AttributeReference
                Dim AttributeCollection As AttributeCollection
                Dim Entity As Entity
                If IsNothing(Document) = True Then
                    Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Else
                    If ActiveDocument = True Then
                        If Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name.Equals(Document.Name) = False Then
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Document
                        End If
                    End If
                End If
                Editor = Document.Editor
                If IsNothing(Transaction) = False Then
                    Database = New Database(False, True)
                    Database.ReadDwgFile(Path.FullName, FileOpenMode.OpenForReadAndReadShare, True, Nothing)
                    If BlockName.Trim = "" Then
                        BlockName = SymbolUtilityServices.GetBlockNameFromInsertPathName(Path.FullName)
                    End If
                    ObjectId = Document.Database.Insert(BlockName, Database, True)
                    If IsNothing(ObjectId) = False Then
                        BlockTableRecord = ObjectId.GetObject(OpenMode.ForWrite)
                        BlockTableRecord.Name = BlockName
                        If IsNothing(LayoutName) = True Then
                            BlockTable = Document.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BlockTableRecord = Transaction.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        Else
                            LayoutManager = LayoutManager.Current
                            LayoutId = LayoutManager.GetLayoutId(LayoutName)
                            Layout = Transaction.GetObject(LayoutId, OpenMode.ForRead)
                            BlockTableRecord = Transaction.GetObject(Layout.BlockTableRecordId, OpenMode.ForWrite)
                        End If
                        BlockReference = New BlockReference(Position, ObjectId)
                        BlockReference.ScaleFactors = ScaleFactors
                        If IsNothing(Layer) = False Then
                            BlockReference.Layer = Layer
                        End If
                        BlockReference.Rotation = Rotation
                        BlockTableRecord.AppendEntity(BlockReference)
                        BlockReference.SetDatabaseDefaults()
                        Transaction.AddNewlyCreatedDBObject(BlockReference, True)
                        BlockTableRecord = Transaction.GetObject(BlockReference.BlockTableRecord.ConvertToRedirectedId, OpenMode.ForRead)
                        If BlockTableRecord.HasAttributeDefinitions = True Then
                            AttributeCollection = BlockReference.AttributeCollection
                            For Each ObjectId2 As ObjectId In BlockTableRecord
                                Entity = ObjectId2.GetObject(OpenMode.ForRead)
                                If Entity.GetType.Name = "AttributeDefinition" Then
                                    AttributeDefinition = Entity
                                    If IsNothing(AttributeDefinition) = False Then
                                        AttributeReference = New AttributeReference()
                                        With AttributeReference
                                            .SetPropertiesFrom(AttributeDefinition)
                                            .Visible = AttributeDefinition.Visible
                                            .SetAttributeFromBlock(AttributeDefinition, BlockReference.BlockTransform)
                                            .HorizontalMode = AttributeDefinition.HorizontalMode
                                            .VerticalMode = AttributeDefinition.VerticalMode
                                            .Rotation = AttributeDefinition.Rotation
                                            .TextStyleId = AttributeDefinition.TextStyleId
                                            .Position = AttributeDefinition.Position + Position.GetAsVector()
                                            .Tag = AttributeDefinition.Tag
                                            .FieldLength = AttributeDefinition.FieldLength
                                            .TextString = AttributeDefinition.TextString
                                            .AdjustAlignment(Database)
                                        End With
                                        AttributeCollection.AppendAttribute(AttributeReference)
                                        Transaction.AddNewlyCreatedDBObject(AttributeReference, True)
                                    End If
                                End If
                            Next
                        End If
                    End If
                Else
                    Using Document.Database
                        Using Editor.Document.LockDocument
                            Transaction = Document.Database.TransactionManager.StartTransaction()
                            Using Transaction
                                Try
                                    BlockReference = Insert(Path, Position, ScaleFactors, Rotation, BlockName, Layer, LayoutName, Document, Transaction, ActiveDocument)
                                    Transaction.Commit()
                                Catch
                                    Transaction.Abort()
                                    Exit Try
                                End Try
                            End Using
                        End Using
                    End Using
                End If
                Return BlockReference
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Insert 
        ''' </summary>
        ''' <param name="Path">Nome e caminho do arquivo a ser inserido</param>
        ''' <param name="Position">Ponto de inserção</param>
        ''' <param name="ScaleFactors">Escala</param>
        ''' <param name="Rotation">Rotação</param>
        ''' <param name="BlockName">Nome do bloco</param>
        ''' <param name="Document">Documento</param>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="ActiveDocument">Determina se o documento será ativado</param>
        ''' <param name="LayoutName">Nome do layout onde será inserido o bloco</param>
        ''' <param name="Layer">Camada</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Insert2(Path As FileInfo, Position As Point3d, ScaleFactors As Scale3d, Optional Rotation As Double = 0, Optional BlockName As String = "", Optional Layer As Object = Nothing, Optional LayoutName As Object = Nothing, Optional Document As Document = Nothing, Optional Transaction As Transaction = Nothing, Optional ActiveDocument As Boolean = True) As BlockReference
            Try
                Dim Editor As Editor
                Dim Database As Database
                Dim ObjectId As ObjectId
                Dim BlockTableRecord As BlockTableRecord = Nothing
                Dim BlockReference As BlockReference = Nothing
                Dim LayoutManager As LayoutManager
                Dim LayoutId As ObjectId
                Dim Layout As Layout
                Dim BlockTable As BlockTable
                Dim TypedValueCollection As New List(Of TypedValue)
                Dim AttributeDefinitionCollection As New List(Of AttributeDefinition)
                Dim AttributeDefinition As AttributeDefinition
                Dim AttributeReference As AttributeReference
                Dim AttributeCollection As AttributeCollection
                Dim Entity As Entity
                If IsNothing(Document) = True Then
                    Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Else
                    If ActiveDocument = True Then
                        If Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name.Equals(Document.Name) = False Then
                            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Document
                        End If
                    End If
                End If
                Editor = Document.Editor
                If IsNothing(Transaction) = False Then
                    Database = New Database(False, True)
                    Database.ReadDwgFile(Path.FullName, FileOpenMode.OpenForReadAndReadShare, True, Nothing)
                    If BlockName.Trim = "" Then
                        BlockName = SymbolUtilityServices.GetBlockNameFromInsertPathName(Path.FullName)
                    End If
                    ObjectId = Document.Database.Insert(BlockName, Database, True)
                    If IsNothing(ObjectId) = False Then
                        BlockTableRecord = ObjectId.GetObject(OpenMode.ForWrite)
                        BlockTableRecord.Name = BlockName
                        If IsNothing(LayoutName) = True Then
                            BlockTable = Document.Database.BlockTableId.GetObject(OpenMode.ForRead)
                            BlockTableRecord = Transaction.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        Else
                            LayoutManager = LayoutManager.Current
                            LayoutId = LayoutManager.GetLayoutId(LayoutName)
                            Layout = Transaction.GetObject(LayoutId, OpenMode.ForRead)
                            BlockTableRecord = Transaction.GetObject(Layout.BlockTableRecordId, OpenMode.ForWrite)
                        End If
                        BlockReference = New BlockReference(Position, ObjectId)
                        BlockReference.ScaleFactors = ScaleFactors
                        If IsNothing(Layer) = False Then
                            BlockReference.Layer = Layer
                        End If
                        BlockReference.Rotation = Rotation
                        BlockTableRecord.AppendEntity(BlockReference)
                        BlockReference.SetDatabaseDefaults()
                        Transaction.AddNewlyCreatedDBObject(BlockReference, True)
                        BlockTableRecord = Transaction.GetObject(BlockReference.BlockTableRecord.ConvertToRedirectedId, OpenMode.ForRead)
                        If BlockTableRecord.HasAttributeDefinitions = True Then
                            AttributeCollection = BlockReference.AttributeCollection
                            For Each ObjectId2 As ObjectId In BlockTableRecord
                                Entity = ObjectId2.GetObject(OpenMode.ForRead)
                                If Entity.GetType.Name = "AttributeDefinition" Then
                                    AttributeDefinition = Entity
                                    If IsNothing(AttributeDefinition) = False Then
                                        AttributeReference = New AttributeReference()
                                        With AttributeReference
                                            .SetPropertiesFrom(AttributeDefinition)
                                            .Visible = AttributeDefinition.Visible
                                            .SetAttributeFromBlock(AttributeDefinition, BlockReference.BlockTransform)
                                            .HorizontalMode = AttributeDefinition.HorizontalMode
                                            .VerticalMode = AttributeDefinition.VerticalMode
                                            .Rotation = AttributeDefinition.Rotation
                                            .TextStyleId = AttributeDefinition.TextStyleId
                                            .Position = AttributeDefinition.Position + Position.GetAsVector()
                                            .Tag = AttributeDefinition.Tag
                                            .FieldLength = AttributeDefinition.FieldLength
                                            .TextString = AttributeDefinition.TextString
                                            .AdjustAlignment(Database)
                                        End With
                                        AttributeCollection.AppendAttribute(AttributeReference)
                                        Transaction.AddNewlyCreatedDBObject(AttributeReference, True)
                                    End If
                                End If
                            Next
                        End If
                    End If
                Else
                    Using Document.Database
                        Using Editor.Document.LockDocument
                            Transaction = Document.Database.TransactionManager.StartTransaction()
                            Using Transaction
                                Try
                                    BlockReference = Insert(Path, Position, ScaleFactors, Rotation, BlockName, Layer, LayoutName, Document, Transaction, ActiveDocument)
                                    Transaction.Commit()
                                Catch
                                    Transaction.Abort()
                                    Exit Try
                                End Try
                            End Using
                        End Using
                    End Using
                End If
                Return BlockReference
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Insert 
        ''' </summary>
        ''' <param name="FileName">Nome e caminho do arquivo a ser inserido</param>
        ''' <param name="Position">Ponto de inserção</param>
        ''' <param name="ScaleFactors">Fatores de escala</param>
        ''' <param name="BlockName">Nome do bloco</param>
        ''' <param name="Explode">Determina se o bloco será explodido durante a inserção</param>
        ''' <param name="Rotation">Rotação (Graus)</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <param name="AngleFormat">Formato informado para a rotação</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Insert(CurrentSpaceId As ObjectId, FileName As FileInfo, Position As Point3d, ScaleFactors As Scale3d, Optional BlockName As String = "", Optional Explode As Boolean = False, Optional Rotation As Double = 0, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = False, Optional AngleFormat As Engine2.Geometry.eAngleFormat = Geometry.eAngleFormat.Degrees, Optional Transaction As Transaction = Nothing) As BlockReference
            Try
                Insert = Nothing
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim ObjectId As ObjectId
                Dim BlockTableRecord As BlockTableRecord = Nothing
                Dim AnnotativeStates As AnnotativeStates
                Dim ObjectContextManager As ObjectContextManager
                Dim ObjectContextCollection As ObjectContextCollection
                Dim DBObject As DBObject
                Dim AttributeDefinition As AttributeDefinition
                Dim AttributeReference As AttributeReference
                If IsNothing(Transaction) = False Then
                    Using Database As New Database(False, True)
                        Database.ReadDwgFile(FileName.FullName, FileShare.Read, True, Nothing)
                        If BlockName.Trim = "" Then
                            BlockName = SymbolUtilityServices.GetBlockNameFromInsertPathName(FileName.Name)
                        End If
                        ObjectId = Document.Database.Insert(BlockName, Database, True)
                        If IsNothing(ObjectId) = False Then
                            If Database.AnnotativeDwg = True Then
                                BlockTableRecord = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                                BlockTableRecord.Annotative = AnnotativeStates.True
                            End If
                            BlockTableRecord = Transaction.GetObject(CurrentSpaceId, OpenMode.ForWrite)
                            Insert = New BlockReference(Position, ObjectId)
                            If IsNothing(Layer) = False Then
                                Insert.Layer = Layer
                            End If
                            If UseUCS = True Then
                                Insert.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                            If AngleFormat = Geometry.eAngleFormat.Degrees Then
                                Insert.Rotation = Engine2.Geometry.DegreeToRadian(Rotation)
                            Else
                                Insert.Rotation = Rotation
                            End If
                            Insert.ScaleFactors = ScaleFactors
                            BlockTableRecord.AppendEntity(Insert)
                            Insert.SetDatabaseDefaults()
                            AnnotativeStates = Insert.Annotative
                            If AnnotativeStates = Autodesk.AutoCAD.DatabaseServices.AnnotativeStates.True Then
                                ObjectContextManager = Database.ObjectContextManager
                                ObjectContextCollection = ObjectContextManager.GetContextCollection("ACDB_ANNOTATIONSCALES")
                                ObjectContexts.AddContext(Insert, ObjectContextCollection.CurrentContext)
                            End If
                            Transaction.AddNewlyCreatedDBObject(Insert, True)
                            BlockTableRecord = Transaction.GetObject(Insert.BlockTableRecord, OpenMode.ForRead)
                            If BlockTableRecord.HasAttributeDefinitions = True Then
                                For Each objID As ObjectId In BlockTableRecord
                                    DBObject = Transaction.GetObject(objID, OpenMode.ForRead)
                                    If DBObject.GetType.Name = "AttributeDefinition" Then
                                        AttributeDefinition = DBObject
                                        If AttributeDefinition.Constant = False Then
                                            AttributeReference = New AttributeReference
                                            Using AttributeReference
                                                AttributeReference.SetAttributeFromBlock(AttributeDefinition, Insert.BlockTransform)
                                                AttributeReference.Position = AttributeDefinition.Position.TransformBy(Insert.BlockTransform)
                                                AttributeReference.TextString = AttributeDefinition.TextString
                                                AttributeReference.Tag = AttributeDefinition.Tag
                                                Insert.AttributeCollection.AppendAttribute(AttributeReference)
                                                Transaction.AddNewlyCreatedDBObject(AttributeReference, True)
                                            End Using
                                        End If
                                    End If
                                Next
                            End If
                            If Explode = True Then
                                Insert.ExplodeToOwnerSpace()
                                Insert.Erase()
                            End If
                        End If
                    End Using
                Else
                    Using DocumentLock As DocumentLock = Document.LockDocument
                        Transaction = Document.TransactionManager.StartTransaction
                        Using Transaction
                            Try
                                Insert = Insert(CurrentSpaceId, FileName, Position, ScaleFactors, BlockName, Explode, Rotation, Layer, UseUCS, AngleFormat, Transaction)
                                Transaction.Commit()
                            Catch
                                Transaction.Abort()
                                Exit Try
                            End Try
                        End Using
                    End Using
                End If
                Return Insert
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Insert 
        ''' </summary>
        ''' <param name="FileName">Nome e caminho do arquivo a ser inserido</param>
        ''' <returns>Size</returns>
        ''' <remarks></remarks>
        Public Shared Function Size(FileName As String) As SizeF
            Try
                Size = New SizeF(0, 0)
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim ObjectId As ObjectId
                Dim BlockTableRecord As BlockTableRecord = Nothing
                Dim BlockReference As BlockReference = Nothing
                Dim Extents3d As Extents3d
                Dim Width As Double = 0
                Dim Height As Double = 0
                Using Database As New Database(False, True)
                    Database.ReadDwgFile(FileName, FileShare.Read, True, Nothing)
                    Using DocumentLock As DocumentLock = Document.LockDocument()
                        Using Transaction As Transaction = Document.TransactionManager.StartTransaction
                            ObjectId = Document.Database.Insert(SymbolUtilityServices.GetBlockNameFromInsertPathName(FileName), Database, True)
                            If IsNothing(ObjectId) = False Then
                                BlockTableRecord = Transaction.GetObject(Document.Database.CurrentSpaceId, OpenMode.ForWrite)
                                BlockReference = New BlockReference(New Point3d(0, 0, 0), ObjectId)
                                Extents3d = BlockReference.GeometricExtents
                                Width = New Point2d(Extents3d.MinPoint.X, 0).GetDistanceTo(New Point2d(Extents3d.MaxPoint.X, 0))
                                Height = New Point2d(0, Extents3d.MaxPoint.Y).GetDistanceTo(New Point2d(0, Extents3d.MinPoint.Y))
                                Size = New SizeF(Width, Height)
                                Transaction.Abort()
                            Else
                                Size = Nothing
                            End If
                        End Using
                    End Using
                End Using
                Return Size
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Atualiza o atributo
        ''' </summary>
        ''' <param name="BlockReference">Bloco</param>
        ''' <param name="Tag">Tag</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function SetAttribute(BlockReference As BlockReference, Tag As String, Value As String, Optional Transaction As Transaction = Nothing) As Boolean
            Dim Close As Boolean = False
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            Dim Entity As Entity
            Dim DBObject As DBObject
            Dim AttributeReference As AttributeReference
            If IsNothing(Transaction) = False Then
                BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                BlockTableRecord = Transaction.GetObject(BlockTable(BlockReference.Name), OpenMode.ForRead)
                For Each ObjectId As ObjectId In BlockTableRecord
                    Entity = TryCast(Transaction.GetObject(ObjectId, OpenMode.ForRead), Entity)
                    If IsNothing(Entity) = False Then
                        For Each ObjectId2 As ObjectId In BlockReference.AttributeCollection
                            DBObject = Transaction.GetObject(ObjectId2, OpenMode.ForRead)
                            AttributeReference = TryCast(DBObject, AttributeReference)
                            If IsNothing(AttributeReference) = False Then
                                If AttributeReference.Tag.Trim.ToUpper = Tag.Trim.ToUpper Then
                                    AttributeReference.UpgradeOpen()
                                    AttributeReference.TextString = Value
                                    AttributeReference.DowngradeOpen()
                                    Close = True
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                    If Close = True Then
                        Exit For
                    End If
                Next
            Else
                Using Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction()
                    Using Transaction
                        SetAttribute(BlockReference, Tag, Value, Transaction)
                        Transaction.Commit()
                    End Using
                End Using
            End If
            Return True
        End Function

        ''' <summary>
        ''' Lê a tag de um atributo em um bloco
        ''' </summary>
        ''' <param name="BlockReference">Bloco</param>
        ''' <param name="Tag">Tag</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ReadAttribute(BlockReference As BlockReference, Tag As String, Optional Transaction As Transaction = Nothing) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim BlockTableRecord As BlockTableRecord
            Dim AttributeCollection As AttributeCollection
            Dim AttributeReference As AttributeReference
            If IsNothing(Transaction) = False Then
                BlockReference = Transaction.GetObject(BlockReference.ObjectId, OpenMode.ForRead)
                BlockTableRecord = Transaction.GetObject(BlockReference.BlockTableRecord, OpenMode.ForRead)
                If BlockTableRecord.HasAttributeDefinitions = True Then
                    AttributeCollection = BlockReference.AttributeCollection
                    For Each ObjectId As ObjectId In AttributeCollection
                        AttributeReference = Transaction.GetObject(ObjectId, OpenMode.ForRead)
                        If AttributeReference.Tag.ToUpper.Trim = Tag.ToUpper.Trim Then
                            Return AttributeReference.TextString
                        End If
                    Next
                    Return Nothing
                Else
                    Return Nothing
                End If
            Else
                Using Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction()
                    Using Transaction
                        Return ReadAttribute(BlockReference, Tag, Transaction)
                        Transaction.Abort()
                    End Using
                End Using
            End If
        End Function

        ''' <summary>
        ''' Desenha bloco
        ''' </summary>
        ''' <param name="BlockName">Nome do bloco</param>
        ''' <param name="EntityCollection">Entidades para composição do bloco</param>
        ''' <param name="Position">Posição</param>
        ''' <param name="ScaleFactors">Fatores de escala</param>
        ''' <param name="Rotation">Rotação</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="Explodable">Determina se o bloco pode ser explodido</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <param name="AngleFormat">Formato informado para a rotação</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>BlockReference</returns>
        ''' <remarks></remarks>
        Public Shared Function Create(BlockName As String, EntityCollection As ObjectIdCollection, Position As Point3d, ScaleFactors As Scale3d, Optional Rotation As Double = 0, Optional Layer As Object = Nothing, Optional Explodable As Boolean = True, Optional UseUCS As Boolean = False, Optional AngleFormat As Engine2.Geometry.eAngleFormat = Geometry.eAngleFormat.Degrees, Optional Transaction As Transaction = Nothing) As BlockReference
            Create = Nothing
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            Dim WorkSpace As BlockTableRecord
            Dim BlockID As ObjectId
            If IsNothing(Transaction) = False Then
                SymbolUtilityServices.ValidateSymbolName(BlockName, False)
                BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForWrite)
                WorkSpace = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite, False)
                If BlockTable.Has(BlockName) = False Then
                    BlockTableRecord = New BlockTableRecord
                    BlockTableRecord.Name = BlockName
                    BlockTableRecord.Origin = Position
                    BlockTableRecord.Explodable = Explodable
                    ' BlockTable.UpgradeOpen()
                    BlockID = BlockTable.Add(BlockTableRecord)
                    ' BlockTable.DowngradeOpen()
                    BlockTableRecord.AssumeOwnershipOf(EntityCollection)
                    Transaction.AddNewlyCreatedDBObject(BlockTableRecord, True)
                    Create = New BlockReference(Position, BlockID)
                    If IsNothing(Layer) = False Then
                        Create.Layer = Layer
                    End If
                    If UseUCS = True Then
                        Create.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    If AngleFormat = Geometry.eAngleFormat.Degrees Then
                        Create.Rotation = Engine2.Geometry.DegreeToRadian(Rotation)
                    Else
                        Create.Rotation = Rotation
                    End If
                    Create.ScaleFactors = ScaleFactors
                    Create.SetDatabaseDefaults()
                    WorkSpace.AppendEntity(Create)
                    Transaction.AddNewlyCreatedDBObject(Create, True)
                Else
                    Throw New System.Exception("O bloco '" & BlockName & "' já existe.")
                End If
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction
                    Using Transaction
                        Try
                            Create = Create(BlockName, EntityCollection, Position, ScaleFactors, Rotation, Layer, Explodable, UseUCS, AngleFormat, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
            End If
            Return Create
        End Function

        ''' <summary>
        ''' Limpa o bloco do sistema
        ''' </summary>
        '''  <param name="BlockName">Nome do bloco</param>
        ''' <param name="Transaction">Transação</param>
        ''' <remarks></remarks>
        Public Shared Sub PurgeBlock(BlockName As String, Optional Transaction As Transaction = Nothing)
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim BlockTable As BlockTable
                Dim BlockTableRecord As BlockTableRecord
                Dim BlockID As ObjectId
                Dim ObjectIdCollection As New ObjectIdCollection
                If IsNothing(Transaction) = False Then
                    BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                    If BlockTable.Has(BlockName) = True Then
                        BlockTableRecord = Transaction.GetObject(BlockTable(BlockName), OpenMode.ForRead)
                        BlockID = BlockTable(BlockName)
                        ObjectIdCollection.Add(BlockID)
                        Database.Purge(ObjectIdCollection)
                    End If
                Else
                    Using DocumentLock As DocumentLock = Document.LockDocument
                        Transaction = Document.TransactionManager.StartTransaction
                        Using Transaction
                            Try
                                PurgeBlock(BlockName, Transaction)
                                Transaction.Commit()
                            Catch
                                Transaction.Abort()
                                Exit Try
                            End Try
                        End Using
                    End Using
                End If
            Catch
                Exit Try
            End Try
        End Sub

        ''' <summary>
        ''' Redefine um bloco
        ''' </summary>
        ''' <param name="CurrentBlockName">Nome do bloco corrente (inserido)</param>
        ''' <param name="ReplacementFilenameBlock">Caminho para o novo bloco</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Redefine(CurrentBlockName As String, ReplacementFilenameBlock As String, Optional Transaction As Transaction = Nothing) As BlockReference
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim ExternalDatabase As Database
            Dim BlockTable As BlockTable
            Dim ObjectId As ObjectId
            Dim BlockTableRecord As BlockTableRecord
            Dim ObjectIdCollection As ObjectIdCollection
            Dim BlockReference As BlockReference = Nothing
            Dim CreateTransaction As Boolean = False
            ExternalDatabase = New Database(False, True)
            ExternalDatabase.ReadDwgFile(ReplacementFilenameBlock, System.IO.FileShare.Read, True, "")
            Using DocumentLock As DocumentLock = Document.LockDocument()
                If IsNothing(Transaction) = True Then
                    Transaction = Database.TransactionManager.StartTransaction()
                    CreateTransaction = True
                End If
                Using Transaction
                    Try
                        BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead, False, True)
                        ObjectId = Database.Insert(CurrentBlockName, ExternalDatabase, True)
                        If ObjectId <> ObjectId.Null Then
                            BlockTableRecord = Transaction.GetObject(ObjectId, OpenMode.ForRead, False, True)
                            ObjectIdCollection = BlockTableRecord.GetBlockReferenceIds(False, True)
                            For Each Id As ObjectId In ObjectIdCollection
                                BlockReference = Transaction.GetObject(Id, OpenMode.ForWrite, False, True)
                                BlockReference.RecordGraphicsModified(True)
                            Next
                        End If
                        If CreateTransaction = True Then
                            Transaction.Commit()
                        End If
                    Catch
                        If CreateTransaction = True Then
                            Transaction.Abort()
                        End If
                    End Try
                End Using
            End Using
            ExternalDatabase.Dispose()
            Return BlockReference
        End Function

        ''' <summary>
        ''' Retorna a imagem bitmap do bloco
        ''' </summary>
        ''' <param name="BlockName">Nome do bloco</param>
        ''' <returns>Bitmap</returns>
        ''' <remarks></remarks>
        Public Shared Function GetBitmap(BlockName As String) As Bitmap
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            Dim ImageSource As ImageSource
            Using DocumentLock As DocumentLock = Document.LockDocument()
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                    If BlockTable.Has(BlockName) = True Then
                        BlockTableRecord = Transaction.GetObject(BlockTable(BlockName), OpenMode.ForRead)
                        ImageSource = CMLContentSearchPreviews.GetBlockTRThumbnail(BlockTableRecord)
                        Return Engine2.Block.ImageSourceToGDI(ImageSource)
                    Else
                        Return Nothing
                    End If
                    Transaction.Abort()
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Obtem a coleção de nomes de blocos
        ''' </summary>
        ''' <param name="BlockFilter">Filtro</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>ArrayList</returns>
        ''' <remarks></remarks>
        Public Shared Function GetNames(BlockFilter As eBlockFilter, Optional Transaction As Transaction = Nothing) As ArrayList
            GetNames = New ArrayList
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim BlockTableRecord As BlockTableRecord
            Dim BlockTable As BlockTable
            If IsNothing(Transaction) = False Then
                BlockTable = Transaction.GetObject(Document.Database.BlockTableId, OpenMode.ForRead)
                For Each ObjectID As ObjectId In BlockTable
                    BlockTableRecord = Transaction.GetObject(ObjectID, OpenMode.ForWrite)
                    Select Case BlockFilter
                        Case eBlockFilter.All
                            If BlockTableRecord.IsLayout = False And BlockTableRecord.IsAnonymous = False And BlockTableRecord.IsFromOverlayReference = False Then
                                GetNames.Add(BlockTableRecord.Name)
                            End If
                        Case eBlockFilter.Block
                            If BlockTableRecord.IsLayout = False And BlockTableRecord.IsAnonymous = False And BlockTableRecord.IsFromExternalReference = False And BlockTableRecord.IsFromOverlayReference = False Then
                                GetNames.Add(BlockTableRecord.Name)
                            End If
                        Case eBlockFilter.XRef
                            If BlockTableRecord.IsLayout = False And BlockTableRecord.IsAnonymous = False And BlockTableRecord.IsFromExternalReference = True And BlockTableRecord.IsFromOverlayReference = False Then
                                GetNames.Add(BlockTableRecord.Name)
                            End If
                    End Select
                Next
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Transaction = Document.Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            GetNames = GetNames(BlockFilter, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
            End If
            Return GetNames
        End Function

        ''' <summary>
        ''' Converte Stream da imagem em GDI
        ''' </summary>
        ''' <param name="src"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function ImageSourceToGDI(src As System.Windows.Media.Imaging.BitmapSource) As System.Drawing.Image
            Dim BmpBitmapEncoder As BmpBitmapEncoder
            Using MemoryStream = New MemoryStream()
                BmpBitmapEncoder = New System.Windows.Media.Imaging.BmpBitmapEncoder()
                With BmpBitmapEncoder
                    .Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(src))
                    .Save(MemoryStream)
                End With
                MemoryStream.Flush()
                Return System.Drawing.Image.FromStream(MemoryStream)
            End Using
        End Function

        ''' <summary>
        ''' Retorna se um bloco é Xref
        ''' </summary>
        ''' <param name="BlockName">Nome do bloco</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function IsXref(BlockName As String, Optional Transaction As Transaction = Nothing) As Boolean
            IsXref = False
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            If IsNothing(Transaction) = False Then
                BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                If BlockTable.Has(BlockName) = True Then
                    BlockTableRecord = Transaction.GetObject(BlockTable(BlockName), OpenMode.ForRead)
                    IsXref = BlockTableRecord.IsFromExternalReference
                End If
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Transaction = Document.Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            IsXref = IsXref(BlockName, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
                Return IsXref
            End If
        End Function

        ''' <summary>
        ''' Explode bloco no espaço do usuário
        ''' </summary>
        ''' <param name="BlockReference">Bloco</param>
        ''' <param name="AllowedClassNameCollection">Nomes das classes a serem consideradas para explosão (Outras serão excluídas).</param>
        ''' <param name="OriginalErase">Determina se a entidade original deve ser excluída</param>
        ''' <returns>Coleção de itens filtrados</returns>
        ''' <remarks></remarks>
        Public Shared Function BlockExplodeToOwnerSpace(BlockReference As BlockReference, Optional AllowedClassNameCollection As ArrayList = Nothing, Optional OriginalErase As Boolean = True) As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection
            BlockExplodeToOwnerSpace = New Autodesk.AutoCAD.DatabaseServices.DBObjectCollection
            Dim ExcludeItens As New Autodesk.AutoCAD.DatabaseServices.DBObjectCollection
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Handler As ObjectEventHandler = Function(sender As Object, e As ObjectEventArgs)
                                                    If IsNothing(AllowedClassNameCollection) = True Then
                                                        BlockExplodeToOwnerSpace.Add(e.DBObject)
                                                    ElseIf AllowedClassNameCollection.Contains(e.DBObject.GetType.Name) = True Then
                                                        BlockExplodeToOwnerSpace.Add(e.DBObject)
                                                    Else
                                                        If Engine2.EntityInteration.IsEntity(e.DBObject) = True Then
                                                            ExcludeItens.Add(e.DBObject)
                                                        End If
                                                    End If
                                                    Return Nothing
                                                End Function
            Using DocumentLock As DocumentLock = Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                    BlockReference = Transaction.GetObject(BlockReference.ObjectId, OpenMode.ForWrite)
                    AddHandler Database.ObjectAppended, Handler
                    BlockReference.ExplodeToOwnerSpace()
                    RemoveHandler Database.ObjectAppended, Handler
                    For Each DBObject As DBObject In ExcludeItens
                        If DBObject.IsErased = False Then
                            DBObject = Transaction.GetObject(DBObject.ObjectId, OpenMode.ForWrite)
                            DBObject.Erase()
                        End If
                    Next
                    If OriginalErase = True Then
                        If BlockReference.IsErased = False Then
                            BlockReference.Erase()
                        End If
                    End If
                    Transaction.Commit()
                End Using
            End Using
            Return BlockExplodeToOwnerSpace
        End Function

        ''' <summary>
        ''' Explode bloco no espaço do usuário
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="BlockReference">Bloco</param>
        ''' <param name="AllowedClassNameCollection">Nomes das classes a serem consideradas para explosão (Outras serão excluídas).</param>
        ''' <param name="OriginalErase">Determina se a entidade original deve ser excluída</param>
        ''' <returns>Coleção de itens filtrados</returns>
        ''' <remarks></remarks>
        Public Shared Function BlockExplodeToOwnerSpace(Transaction As Transaction, BlockReference As BlockReference, Optional AllowedClassNameCollection As ArrayList = Nothing, Optional OriginalErase As Boolean = True) As Autodesk.AutoCAD.DatabaseServices.DBObjectCollection
            BlockExplodeToOwnerSpace = New Autodesk.AutoCAD.DatabaseServices.DBObjectCollection
            Dim ExcludeItens As New Autodesk.AutoCAD.DatabaseServices.DBObjectCollection
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Handler As ObjectEventHandler = Function(sender As Object, e As ObjectEventArgs)
                                                    If IsNothing(AllowedClassNameCollection) = True Then
                                                        BlockExplodeToOwnerSpace.Add(e.DBObject)
                                                    ElseIf AllowedClassNameCollection.Contains(e.DBObject.GetType.Name) = True Then
                                                        BlockExplodeToOwnerSpace.Add(e.DBObject)
                                                    Else
                                                        If Engine2.EntityInteration.IsEntity(e.DBObject) = True Then
                                                            ExcludeItens.Add(e.DBObject)
                                                        End If
                                                    End If
                                                    Return Nothing
                                                End Function
            BlockReference = Transaction.GetObject(BlockReference.ObjectId, OpenMode.ForWrite)
            AddHandler Database.ObjectAppended, Handler
            BlockReference.ExplodeToOwnerSpace()
            RemoveHandler Database.ObjectAppended, Handler
            For Each DBObject As DBObject In ExcludeItens
                Try
                    If DBObject.IsErased = False Then
                        DBObject = Transaction.GetObject(DBObject.ObjectId, OpenMode.ForWrite)
                        DBObject.Erase()
                    End If
                Catch
                    Exit Try
                End Try
            Next
            If OriginalErase = True Then
                If BlockReference.IsErased = False Then
                    BlockReference.Erase()
                End If
            End If
            Return BlockExplodeToOwnerSpace
        End Function

        ''' <summary>
        ''' Obtem a coleção de blocos
        ''' </summary>
        ''' <param name="BlockFilter">Filtro</param>
        ''' <param name="Transaction">Transaction</param>
        ''' <returns>ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function GetBlockTableRecordIds(BlockFilter As eBlockFilter, Optional Transaction As Transaction = Nothing) As ObjectIdCollection
            GetBlockTableRecordIds = New ObjectIdCollection
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim BlockTableRecord As BlockTableRecord
            Dim BlockTable As BlockTable
            If IsNothing(Transaction) = False Then
                BlockTable = Transaction.GetObject(Document.Database.BlockTableId, OpenMode.ForRead)
                For Each ObjectID As ObjectId In BlockTable
                    BlockTableRecord = Transaction.GetObject(ObjectID, OpenMode.ForWrite)
                    Select Case BlockFilter
                        Case eBlockFilter.All
                            If BlockTableRecord.IsLayout = False And BlockTableRecord.IsAnonymous = False And BlockTableRecord.IsFromOverlayReference = False Then
                                GetBlockTableRecordIds.Add(BlockTableRecord.ObjectId)
                            End If
                        Case eBlockFilter.Block
                            If BlockTableRecord.IsLayout = False And BlockTableRecord.IsAnonymous = False And BlockTableRecord.IsFromExternalReference = False And BlockTableRecord.IsFromOverlayReference = False Then
                                GetBlockTableRecordIds.Add(BlockTableRecord.ObjectId)
                            End If
                        Case eBlockFilter.XRef
                            If BlockTableRecord.IsLayout = False And BlockTableRecord.IsAnonymous = False And BlockTableRecord.IsFromExternalReference = True And BlockTableRecord.IsFromOverlayReference = False Then
                                GetBlockTableRecordIds.Add(BlockTableRecord.ObjectId)
                            End If
                    End Select
                Next
            Else
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Transaction = Document.Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            GetBlockTableRecordIds = GetBlockTableRecordIds(BlockFilter, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
            End If
            Return GetBlockTableRecordIds
        End Function

        ''' <summary>
        ''' Obtem o BlockTableRecord de um bloco
        ''' </summary>
        ''' <param name="Transaction">Transaction</param>
        ''' <param name="BlockReference">BlockReference</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetBlockTableRecord(BlockReference As BlockReference, Optional Transaction As Transaction = Nothing) As BlockTableRecord
            Try
                GetBlockTableRecord = Nothing
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim BlockTable As BlockTable
                If IsNothing(Transaction) = False Then
                    BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                    If BlockTable.Has(BlockReference.Name) = True Then
                        GetBlockTableRecord = Transaction.GetObject(BlockTable(BlockReference.Name), OpenMode.ForRead)
                    End If
                Else
                    Using DocumentLock As DocumentLock = Document.LockDocument
                        Transaction = Document.TransactionManager.StartTransaction
                        Using Transaction
                            Try
                                GetBlockTableRecord = GetBlockTableRecord(BlockReference, Transaction)
                                Transaction.Commit()
                            Catch
                                Transaction.Abort()
                                Exit Try
                            End Try
                        End Using
                    End Using
                End If
                Return GetBlockTableRecord
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Clona um bloco (Não utilizar em XRef)
        ''' </summary>
        ''' <param name="BlockForCopy">Bloco para cópia</param>
        ''' <param name="BlockName">Nome do bloco</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DeepCloneBlock(BlockForCopy As BlockReference, BlockName As String, Optional Transaction As Transaction = Nothing) As BlockTableRecord
            Try
                DeepCloneBlock = Nothing
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim Editor As Editor = Document.Editor
                Dim BlockTable As BlockTable
                Dim BlockReference As BlockReference
                Dim BlockTableRecord As BlockTableRecord
                Dim NewBlockTableRecord As BlockTableRecord
                Dim IdMapping As IdMapping
                If IsNothing(Transaction) = False Then
                    BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForWrite)
                    BlockReference = Transaction.GetObject(BlockForCopy.ObjectId, OpenMode.ForRead)
                    BlockTableRecord = Transaction.GetObject(BlockReference.BlockTableRecord, OpenMode.ForRead)
                    If BlockTable.Has(BlockName) = False Then
                        IdMapping = New IdMapping()
                        NewBlockTableRecord = New BlockTableRecord
                        NewBlockTableRecord = BlockTableRecord.DeepClone(BlockTable, IdMapping, False)
                        NewBlockTableRecord.Name = BlockName
                        Transaction.AddNewlyCreatedDBObject(NewBlockTableRecord, True)
                    Else
                        DeepCloneBlock = Transaction.GetObject(BlockTable(BlockName), OpenMode.ForRead)
                    End If
                Else
                    Using Editor.Document.LockDocument
                        Transaction = Database.TransactionManager.StartTransaction
                        Using Transaction
                            Try
                                DeepCloneBlock = DeepCloneBlock(BlockForCopy, BlockName, Transaction)
                                Transaction.Commit()
                            Catch
                                Transaction.Abort()
                                Exit Try
                            End Try
                        End Using
                    End Using
                End If
                Return DeepCloneBlock
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retorna o bloco de nível mais elevado com base na seleção de um item podendo ser o próprio XRef ou o nível imediatamente a baixo
        ''' </summary>
        ''' <param name="PromptNestedEntityResult">PromptNestedEntityResult</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>BlockReference</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTopLevelBlock(PromptNestedEntityResult As PromptNestedEntityResult, Optional Transaction As Transaction = Nothing) As BlockReference
            Try
                GetTopLevelBlock = Nothing
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim Editor As Editor = Document.Editor
                Dim Entity As Entity
                Dim ObjectIds() As ObjectId
                If IsNothing(Transaction) = False Then
                    Entity = PromptNestedEntityResult.ObjectId.ToEntity(Transaction)
                    ObjectIds = PromptNestedEntityResult.GetContainers
                    For Each ObjectId As ObjectId In ObjectIds
                        Entity = ObjectId.ToEntity(Transaction)
                        If Entity.BlockName.Equals("*Model_Space") = True Then
                            GetTopLevelBlock = Entity
                            Exit For
                        End If
                    Next
                Else
                    Using Editor.Document.LockDocument
                        Transaction = Database.TransactionManager.StartTransaction
                        Using Transaction
                            Try
                                GetTopLevelBlock = GetTopLevelBlock(PromptNestedEntityResult, Transaction)
                                Transaction.Commit()
                            Catch
                                Transaction.Abort()
                                Exit Try
                            End Try
                        End Using
                    End Using
                End If
                Return GetTopLevelBlock
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem o caminho do XRef
        ''' </summary>
        ''' <param name="BlockReference"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetXrefFileInfo(BlockReference As BlockReference) As FileInfo
            Try
                GetXrefFileInfo = Nothing
                Dim BlockTableRecord As BlockTableRecord = DirectCast(BlockReference.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)
                If BlockTableRecord.IsFromExternalReference OrElse BlockTableRecord.IsFromOverlayReference Then
                    If Not BlockTableRecord.IsUnloaded AndAlso BlockTableRecord.XrefStatus = XrefStatus.Resolved Then
                        Using Database As Database = BlockTableRecord.GetXrefDatabase(False)
                            If Database IsNot Nothing Then
                                GetXrefFileInfo = New FileInfo(Database.Filename)
                            End If
                        End Using
                    Else
                        GetXrefFileInfo = New FileInfo(HostApplicationServices.Current.FindFile(BlockTableRecord.PathName, BlockTableRecord.Database, FindFileHint.XRefDrawing))
                    End If
                End If
                Return GetXrefFileInfo
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Importa um bloco de um XRef
        ''' </summary>
        ''' <param name="XRefBlock">Bloco XRef</param>
        ''' <param name="CopyBlockName">Nome do bloco a ser copiado</param>
        ''' <param name="NewBlockTableRecordName">Nome para o novo BlockTableRecord</param>
        ''' <param name="Explodable">Determina se o bloco pode ser explodido</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ImportBlock(XRefBlock As BlockReference, CopyBlockName As String, Optional NewBlockTableRecordName As Object = Nothing, Optional Explodable As Boolean = True, Optional Transaction As Transaction = Nothing) As BlockTableRecord
            ImportBlock = Nothing
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim LocalDatabase As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim LocalBlockTableRecord As BlockTableRecord
            Dim RemoteBlockTableRecord As BlockTableRecord
            Dim Path As String = ""
            Dim blockIds As New ObjectIdCollection
            Dim LocalBlockTable As BlockTable
            Dim RemoteBlockTable As BlockTable
            Dim FileInfo As FileInfo
            Try
                If IsNothing(Transaction) = False Then
                    If CopyBlockName.Contains("|") = True Then
                        CopyBlockName = CopyBlockName.Split("|")(1)
                    End If
                    LocalBlockTable = Transaction.GetObject(LocalDatabase.BlockTableId, OpenMode.ForRead, False)
                    If IsNothing(NewBlockTableRecordName) = True Then
                        If LocalBlockTable.Has(CopyBlockName) = True Then
                            ImportBlock = Transaction.GetObject(LocalBlockTable(CopyBlockName), OpenMode.ForRead)
                        End If
                    Else
                        If LocalBlockTable.Has(NewBlockTableRecordName) = True Then
                            ImportBlock = Transaction.GetObject(LocalBlockTable(NewBlockTableRecordName), OpenMode.ForRead)
                        End If
                    End If
                    If IsNothing(ImportBlock) = True Then
                        LocalBlockTableRecord = Transaction.GetObject(XRefBlock.BlockTableRecord, OpenMode.ForRead)
                        FileInfo = GetXrefFileInfo(XRefBlock)
                        Path = FileInfo.FullName
                        If Engine2.CurrentUserAccess.HasAccess(New FileInfo(Path), Security.AccessControl.FileSystemRights.ReadData) = False Then
                            Throw New System.Exception("O caminho '" & Path & "' não esta acessível.")
                        Else
                            Using RemoteDatabase As New Database(False, True)
                                RemoteDatabase.ReadDwgFile(Path, System.IO.FileShare.Read, True, "")
                                Using RemoteTransaction As Transaction = RemoteDatabase.TransactionManager.StartTransaction
                                    RemoteBlockTable = RemoteTransaction.GetObject(RemoteDatabase.BlockTableId, OpenMode.ForRead, False)
                                    For Each ObjectId As ObjectId In RemoteBlockTable
                                        RemoteBlockTableRecord = RemoteTransaction.GetObject(ObjectId, OpenMode.ForRead, False)
                                        If RemoteBlockTableRecord.IsAnonymous = False AndAlso RemoteBlockTableRecord.IsLayout = False Then
                                            If RemoteBlockTableRecord.Name.Equals(CopyBlockName) = True Then
                                                blockIds.Add(ObjectId)
                                                RemoteDatabase.WblockCloneObjects(blockIds, LocalDatabase.BlockTableId, New IdMapping(), DuplicateRecordCloning.Replace, False)
                                                ImportBlock = Transaction.GetObject(LocalBlockTable(CopyBlockName), OpenMode.ForWrite)
                                                ImportBlock.Explodable = Explodable
                                                If IsNothing(NewBlockTableRecordName) = False Then
                                                    ImportBlock.Name = NewBlockTableRecordName
                                                End If
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End Using
                                RemoteDatabase.Dispose()
                            End Using

                        End If
                    End If
                Else
                    Using Editor.Document.LockDocument
                        Transaction = LocalDatabase.TransactionManager.StartTransaction
                        Using Transaction
                            Try
                                ImportBlock = ImportBlock(XRefBlock, CopyBlockName, NewBlockTableRecordName, Explodable, Transaction)
                                Transaction.Commit()
                            Catch
                                Transaction.Abort()
                                Exit Try
                            End Try
                        End Using
                    End Using
                End If
                Return ImportBlock
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem o nome real do bloco
        ''' </summary>
        ''' <param name="BlockReference">Bloco</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetRealName(BlockReference As BlockReference, Optional Transaction As Transaction = Nothing) As String
            Try
                GetRealName = ""
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim Editor As Editor = Document.Editor
                Dim BlockTableRecord As BlockTableRecord
                If IsNothing(Transaction) = False Then
                    If BlockReference.IsDynamicBlock = True Or BlockReference.Name.Contains("*") = True Then
                        Try
                            BlockTableRecord = Transaction.GetObject(BlockReference.DynamicBlockTableRecord, OpenMode.ForRead)
                        Catch
                            BlockTableRecord = Transaction.GetObject(BlockReference.BlockTableRecord, OpenMode.ForRead)
                        End Try
                    Else
                        BlockTableRecord = Transaction.GetObject(BlockReference.BlockTableRecord, OpenMode.ForRead)
                    End If
                    GetRealName = BlockTableRecord.Name
                Else
                    Using Editor.Document.LockDocument
                        Transaction = Database.TransactionManager.StartTransaction
                        Using Transaction
                            Try
                                GetRealName = GetRealName(BlockReference, Transaction)
                                Transaction.Commit()
                            Catch
                                Transaction.Abort()
                                Exit Try
                            End Try
                        End Using
                    End Using
                End If
                Return GetRealName
            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Obtem os blocos do desenho
        ''' </summary>
        ''' <param name="BlockFilter">Filtro</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>ObjectIdCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function GetBlockIds(BlockFilter As eBlockFilter, Optional Transaction As Transaction = Nothing) As ObjectIdCollection
            GetBlockIds = New ObjectIdCollection
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            If IsNothing(Transaction) = False Then
                BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                For Each ObjectId As ObjectId In BlockTable
                    BlockTableRecord = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                    Select Case BlockFilter
                        Case eBlockFilter.All
                            If BlockTableRecord.IsLayout = False And BlockTableRecord.IsAnonymous = False And BlockTableRecord.IsFromOverlayReference = False Then
                                For Each ObjectId2 As ObjectId In BlockTableRecord.GetBlockReferenceIds(False, True)
                                    GetBlockIds.Add(ObjectId2)
                                Next
                            End If
                        Case eBlockFilter.Block
                            If BlockTableRecord.IsLayout = False And BlockTableRecord.IsAnonymous = False And BlockTableRecord.IsFromExternalReference = False And BlockTableRecord.IsFromOverlayReference = False Then
                                For Each ObjectId2 As ObjectId In BlockTableRecord.GetBlockReferenceIds(False, True)
                                    GetBlockIds.Add(ObjectId2)
                                Next
                            End If
                        Case eBlockFilter.XRef
                            If BlockTableRecord.IsLayout = False And BlockTableRecord.IsAnonymous = False And BlockTableRecord.IsFromExternalReference = True And BlockTableRecord.IsFromOverlayReference = False Then
                                For Each ObjectId2 As ObjectId In BlockTableRecord.GetBlockReferenceIds(False, True)
                                    GetBlockIds.Add(ObjectId2)
                                Next
                            End If
                    End Select
                Next
            Else
                Using Editor.Document.LockDocument
                    Transaction = Database.TransactionManager.StartTransaction
                    Using Transaction
                        Try
                            GetBlockIds = GetBlockIds(BlockFilter, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
            End If
            Return GetBlockIds
        End Function

        ''' <summary>
        ''' Recarrega todos os XRefs
        ''' </summary>
        ''' <param name="XRefNames">Coleção de nomes do XRef</param>
        ''' <param name="Transaction">Transação</param>
        ''' <remarks></remarks>
        Public Shared Sub ReloadXrefs(Optional XRefNames As List(Of String) = Nothing, Optional Transaction As Transaction = Nothing)
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim ObjectIdCollection As New ObjectIdCollection
            If IsNothing(Transaction) = False Then
                Dim XrefGraph As XrefGraph = Database.GetHostDwgXrefGraph(False)
                For Index As Integer = 0 To XrefGraph.NumNodes - 1
                    Dim XrefGraphNode As XrefGraphNode = XrefGraph.GetXrefNode(Index)
                    If XrefGraphNode.BlockTableRecordId.IsValid = True Then
                        If IsNothing(XRefNames) = False Then
                            If XRefNames.Contains(XrefGraphNode.Name) = True And XrefGraphNode.XrefStatus = XrefStatus.Resolved Then
                                ObjectIdCollection.Add(XrefGraphNode.BlockTableRecordId)
                            End If
                        Else
                            If XrefGraphNode.XrefStatus = XrefStatus.Resolved Then
                                ObjectIdCollection.Add(XrefGraphNode.BlockTableRecordId)
                            End If
                        End If
                    End If
                Next
                If ObjectIdCollection.Count > 0 Then
                    Database.ReloadXrefs(ObjectIdCollection)
                End If
                Editor.WriteMessage(vbLf & ObjectIdCollection.Count.ToString & " xref(s) recarregado(s)." & vbLf)
            Else
                Using Editor.Document.LockDocument
                    Transaction = Database.TransactionManager.StartTransaction
                    Using Transaction
                        Try
                            ReloadXrefs(XRefNames, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using

            End If
        End Sub


        ''' <summary>
        ''' Seta o valor da propriedade dinâmica
        ''' </summary>
        ''' <param name="BlockReference"></param>
        ''' <param name="Property"></param>
        ''' <param name="Value"></param>
        ''' <param name="Transaction"></param>
        ''' <remarks></remarks>
        Public Shared Function SetDynamicProperty(BlockReference As BlockReference, [Property] As String, Value As Object, Optional Transaction As Transaction = Nothing) As Boolean
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            If IsNothing(Transaction) = False Then
                BlockReference = Transaction.GetObject(BlockReference.Id, OpenMode.ForWrite)
                For Each Prop As DynamicBlockReferenceProperty In BlockReference.DynamicBlockReferencePropertyCollection
                    If Prop.PropertyName.Equals([Property], StringComparison.OrdinalIgnoreCase) = True Then
                        Prop.Value = Value
                        Return True
                        Exit For
                    End If
                Next
            Else
                Using Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            Return SetDynamicProperty(BlockReference, [Property], Value, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                        End Try
                    End Using
                End Using
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Obtem os valores dinâmicos
        ''' </summary>
        ''' <param name="BlockReference"></param>
        ''' <param name="Transaction"></param>
        ''' <remarks></remarks>
        Public Shared Function GetAllDynamicProperty(BlockReference As BlockReference, Optional Transaction As Transaction = Nothing) As Dictionary(Of String, Object)
            GetAllDynamicProperty = New Dictionary(Of String, Object)
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            If IsNothing(Transaction) = False Then
                BlockReference = Transaction.GetObject(BlockReference.Id, OpenMode.ForWrite)
                For Each Prop As DynamicBlockReferenceProperty In BlockReference.DynamicBlockReferencePropertyCollection
                    GetAllDynamicProperty.Add(Prop.PropertyName, Prop.Value)
                Next
                Return GetAllDynamicProperty
            Else
                Using Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            Return GetAllDynamicProperty(BlockReference, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                        End Try
                    End Using
                End Using
            End If
            Return GetAllDynamicProperty
        End Function

        ''' <summary>
        ''' Seta o valor da propriedade dinâmica
        ''' </summary>
        ''' <param name="BlockReference"></param>
        ''' <param name="Property"></param>
        ''' <param name="Transaction"></param>
        ''' <remarks></remarks>
        Public Shared Function GetDynamicProperty(BlockReference As BlockReference, [Property] As String, Optional Transaction As Transaction = Nothing) As Object
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            If IsNothing(Transaction) = False Then
                BlockReference = Transaction.GetObject(BlockReference.Id, OpenMode.ForWrite)
                For Each Prop As DynamicBlockReferenceProperty In BlockReference.DynamicBlockReferencePropertyCollection
                    If Prop.PropertyName.Equals([Property], StringComparison.OrdinalIgnoreCase) = True Then
                        Return Prop.Value
                        Exit For
                    End If
                Next
            Else
                Using Document.LockDocument
                    Transaction = Document.TransactionManager.StartTransaction()
                    Using Transaction
                        Try
                            Return GetDynamicProperty(BlockReference, [Property], Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                        End Try
                    End Using
                End Using
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Obtem o tamanho do bloco
        ''' </summary>
        ''' <param name="BlockReference">Bloco</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns></returns>
        Public Shared Function Size(BlockReference As BlockReference, Optional Transaction As Transaction = Nothing) As SizeF
            Size = Nothing
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim BlockName As String = BlockReference.BlockName
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            Dim Bounds As Extents3d?
            Dim Width As Double = 0
            Dim Height As Double = 0
            Dim Extents3d As Extents3d
            If IsNothing(Transaction) = False Then
                BlockTable = TryCast(Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead), BlockTable)
                If BlockTable.Has(BlockName) = True Then
                    BlockTableRecord = CType(Transaction.GetObject(BlockTable(BlockName), OpenMode.ForRead, False), BlockTableRecord)
                    Bounds = BlockTableRecord.Bounds
                    If Bounds.HasValue Then
                        Extents3d = Bounds.Value
                        Width = Extents3d.MaxPoint.X - Extents3d.MinPoint.X
                        Height = Extents3d.MaxPoint.Y - Extents3d.MinPoint.Y
                    Else
                        BlockReference = New BlockReference(BlockReference.Position, BlockTable(BlockName))
                        Bounds = BlockReference.Bounds
                        Extents3d = Bounds.Value
                        Width = Extents3d.MaxPoint.X - Extents3d.MinPoint.X
                        Height = Extents3d.MaxPoint.Y - Extents3d.MinPoint.Y
                        BlockReference.Dispose()
                    End If
                    Size = New SizeF(Width, Height)
                End If
            Else
                Using Document.LockDocument
                    Transaction = Database.TransactionManager.StartTransaction
                    Using Transaction
                        Try
                            Size = Size(BlockReference, Transaction)
                            Transaction.Commit()
                        Catch
                            Transaction.Abort()
                            Exit Try
                        End Try
                    End Using
                End Using
            End If
            Return Size
        End Function

    End Class

End Namespace
