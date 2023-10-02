Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput


Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.Runtime


'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Classe para manipulação de XRef´s
    ''' </summary>
    ''' <remarks></remarks>
    Public Class XRef

        ''' <summary>
        ''' Inserir Xref
        ''' </summary>
        ''' <param name="Path">Caminho do bloco</param>
        ''' <param name="Position">Posição de inserção</param>
        ''' <param name="BlockName">Nome do bloco (Caso omitido assume o nome do arquivo)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Attach(Path As String, Position As Point3d, Optional BlockName As String = "") As BlockReference
            Try
                Dim BlockReference As BlockReference = Nothing
                Dim BlockTableRecord As BlockTableRecord
                Dim Database As Autodesk.AutoCAD.DatabaseServices.Database
                Database = Application.DocumentManager.MdiActiveDocument.Database
                Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Dim ObjectId As ObjectId
                If BlockName = "" Then
                    BlockName = System.IO.Path.GetFileNameWithoutExtension(Path)
                End If
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        ObjectId = Database.AttachXref(Path, BlockName)
                        If ObjectId.IsNull = False Then
                            BlockReference = New BlockReference(Position, ObjectId)
                            BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                            BlockTableRecord.AppendEntity(BlockReference)
                            Transaction.AddNewlyCreatedDBObject(BlockReference, True)
                            Attach = BlockReference
                        End If
                        Transaction.Commit()
                    End Using
                End Using
                Return BlockReference
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Remover Xref
        ''' </summary>
        ''' <param name="BlockReference">BlockReference</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Detach(BlockReference As BlockReference) As Boolean
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTableRecord As BlockTableRecord
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    If IsNothing(BlockReference) = False Then
                        If BlockReference.IsErased = False Then
                            BlockTableRecord = DirectCast(Transaction.GetObject(BlockReference.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                            Database.DetachXref(BlockReference.BlockTableRecord)
                            Editor.WriteMessage(vbLf & "External reference detached." & vbLf)
                        End If
                    End If
                    Transaction.Commit()
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Remover Xref
        ''' </summary>
        ''' <param name="ObjectId">ID do bloco</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Detach(ObjectId As ObjectId) As Boolean
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockReference As BlockReference
            Dim BlockTableRecord As BlockTableRecord
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    BlockReference = Engine2.ConvertObject.ObjectIDToEntity(ObjectId)
                    If IsNothing(BlockReference) = False Then
                        If BlockReference.IsErased = False Then
                            BlockTableRecord = DirectCast(Transaction.GetObject(BlockReference.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                            Database.DetachXref(BlockReference.BlockTableRecord)
                            Editor.WriteMessage(vbLf & "External reference detached." & vbLf)
                        End If
                    End If
                    Transaction.Commit()
                End Using
            End Using
        End Function

    End Class

End Namespace





