Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.Windows.Media.Media3D

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Insere Mtexto no desenho
    ''' </summary>
    ''' <remarks></remarks>
    Public Class AcadText

        ''' <summary>
        ''' Cria um texto 
        ''' </summary>
        ''' <param name="Point">Ponto</param>
        ''' <param name="Text">Texto</param>
        ''' <param name="TextHeight">Tamanho da fonte</param>
        ''' <param name="TextStyleName">Estilo de texto</param>
        ''' <param name="Attachment">Alinhamento</param>
        ''' <param name="Rotation">Rotação</param>
        ''' <remarks></remarks>
        Public Shared Function Add(Point As Autodesk.AutoCAD.Geometry.Point3d, Text As String, TextHeight As Double, Optional TextStyleName As String = "Standard", Optional Attachment As Autodesk.AutoCAD.DatabaseServices.AttachmentPoint = AttachmentPoint.MiddleCenter, Optional Rotation As Double = 0, Optional LayerName As String = "0") As MText
            Dim MText As MText = Nothing
            Dim acadDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acadDatabase As Database = acadDocument.Database
            Using acadDocument.Editor.Document.LockDocument
                Using Transaction As Transaction = acadDatabase.TransactionManager.StartTransaction()
                    Dim BlockTable As BlockTable = Transaction.GetObject(acadDatabase.BlockTableId, OpenMode.ForRead)
                    Dim BlockTableRecord As BlockTableRecord = Transaction.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    MText = New MText()
                    With MText
                        .SetDatabaseDefaults()
                        .TextStyleId = Engine2.TextStyle.GetTextStyle(TextStyleName).ObjectId
                        .Location = Point
                        .TextHeight = TextHeight
                        .Contents = Text
                        .Rotation = Rotation
                        .Attachment = Attachment
                        .Layer = LayerName
                    End With
                    BlockTableRecord.AppendEntity(MText)
                    Transaction.AddNewlyCreatedDBObject(MText, True)
                    Transaction.Commit()
                    Transaction.Dispose()
                End Using
            End Using
            Return MText
        End Function

        ''' <summary>
        ''' Atualiza um Mtext
        ''' </summary>
        ''' <param name="Entity">Entity</param>
        ''' <param name="Point">Ponto (3D)</param>
        ''' <param name="Text">Texto</param>
        ''' <param name="TextHeight">Tamanho da fonte</param>
        ''' <param name="TextStyleName">Estilo de texto</param>
        ''' <param name="Attachment">Alinhamento</param>
        ''' <param name="Rotation">Rotação</param>
        ''' <remarks></remarks>
        Public Shared Function Edit(Entity As Entity, Optional Point As Object = Nothing, Optional Text As Object = Nothing, Optional TextHeight As Object = Nothing, Optional TextStyleName As Object = Nothing, Optional Attachment As Object = Nothing, Optional Rotation As Object = Nothing, Optional LayerName As Object = Nothing) As MText
            Dim acadDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acadDatabase As Database = acadDocument.Database
            Dim Mtext As MText = Nothing
            Using acadDocument.Editor.Document.LockDocument
                Using Transaction As Transaction = acadDatabase.TransactionManager.StartTransaction()
                    Dim BlockTable As BlockTable = Transaction.GetObject(acadDatabase.BlockTableId, OpenMode.ForRead)
                    Dim BlockTableRecord As BlockTableRecord = Transaction.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Entity = DirectCast(Transaction.GetObject(Entity.Id, OpenMode.ForRead), Entity)
                    Mtext = Entity
                    With Mtext
                        .UpgradeOpen()
                        If IsNothing(TextStyleName) = False Then
                            .TextStyleId = Engine2.TextStyle.GetTextStyle(TextStyleName).ObjectId
                        End If
                        If IsNothing(TextHeight) = False Then
                            .TextHeight = TextHeight
                        End If
                        If IsNothing(Text) = False Then
                            .Contents = Text
                        End If
                        If IsNothing(Rotation) = False Then
                            .Rotation = Rotation
                        End If
                        If IsNothing(Attachment) = False Then
                            .Attachment = Attachment
                        End If
                        If IsNothing(Point) = False Then
                            .Location = Point
                        End If
                        If IsNothing(LayerName) = False Then
                            .Layer = LayerName
                        End If
                        .RecordGraphicsModified(True)
                    End With
                    Transaction.Commit()
                    Transaction.Dispose()
                End Using
            End Using
            Return Mtext
        End Function

        ''' <summary>
        ''' Atualiza um Mtext
        ''' </summary>
        ''' <param name="ObjectID">ObjectID</param>
        ''' <param name="Point">Ponto (3D)</param>
        ''' <param name="Text">Texto</param>
        ''' <param name="TextHeight">Tamanho da fonte</param>
        ''' <param name="TextStyleName">Estilo de texto</param>
        ''' <param name="Attachment">Alinhamento</param>
        ''' <param name="Rotation">Rotação</param>
        ''' <remarks></remarks>
        Public Shared Function Edit(ObjectID As ObjectId, Optional Point As Object = Nothing, Optional Text As Object = Nothing, Optional TextHeight As Object = Nothing, Optional TextStyleName As Object = Nothing, Optional Attachment As Object = Nothing, Optional Rotation As Object = Nothing, Optional LayerName As Object = Nothing) As MText
            Dim acadDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acadDatabase As Database = acadDocument.Database
            Dim Mtext As MText = Nothing
            Using acadDocument.Editor.Document.LockDocument
                Using Transaction As Transaction = acadDatabase.TransactionManager.StartTransaction()
                    Dim BlockTable As BlockTable = Transaction.GetObject(acadDatabase.BlockTableId, OpenMode.ForRead)
                    Dim BlockTableRecord As BlockTableRecord = Transaction.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim Entity As Entity = DirectCast(Transaction.GetObject(ObjectID, OpenMode.ForRead), Entity)
                    Mtext = Entity
                    With Mtext
                        .UpgradeOpen()
                        If IsNothing(TextStyleName) = False Then
                            .TextStyleId = Engine2.TextStyle.GetTextStyle(TextStyleName).ObjectId
                        End If
                        If IsNothing(TextHeight) = False Then
                            .TextHeight = TextHeight
                        End If
                        If IsNothing(Text) = False Then
                            .Contents = Text
                        End If
                        If IsNothing(Rotation) = False Then
                            .Rotation = Rotation
                        End If
                        If IsNothing(Attachment) = False Then
                            .Attachment = Attachment
                        End If
                        If IsNothing(Point) = False Then
                            .Location = Point
                        End If
                        If IsNothing(LayerName) = False Then
                            .Layer = LayerName
                        End If
                        .RecordGraphicsModified(True)
                    End With
                    Transaction.Commit()
                    Transaction.Dispose()
                End Using
            End Using
            Return Mtext
        End Function



    End Class

End Namespace


