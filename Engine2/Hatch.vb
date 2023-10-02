'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Reflection


Namespace Engine2

    ''' <summary>
    ''' Classe para inclusão de hachura na entidade
    ''' </summary>
    ''' <remarks></remarks>
    Public Class HatchEngine

        ''' <summary>
        ''' Adiciona hachura na entidade
        ''' </summary>
        ''' <param name="DBObject">Objeto DataBase</param>
        ''' <param name="HatchName">Nome da hachura</param>
        ''' <param name="HatchScale">Escala</param>
        ''' <param name="Transparency">Transparência</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AddHatch(ByVal DBObject As DBObject, ByVal HatchName As String, ByVal HatchScale As Double, HatchAngle As Double, Optional Transparency As Integer = 0) As Hatch
            Try
                Transparency = Math.Truncate(255 * (100 - Transparency) / 100)
                Dim Hatch As Hatch = Nothing
                Dim acadDocument As Document = Application.DocumentManager.MdiActiveDocument
                Dim acadDatabase As Database = acadDocument.Database
                Using Transaction As Transaction = acadDatabase.TransactionManager.StartTransaction()
                    Dim BlockTable As BlockTable = Transaction.GetObject(acadDatabase.BlockTableId, OpenMode.ForRead)
                    Dim BlockTableRecord As BlockTableRecord = Transaction.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim ObjectIdCollection As ObjectIdCollection = New ObjectIdCollection()
                    ObjectIdCollection.Add(DBObject.ObjectId)
                    Hatch = New Hatch()
                    BlockTableRecord.AppendEntity(Hatch)
                    Transaction.AddNewlyCreatedDBObject(Hatch, True)
                    With Hatch
                        .SetDatabaseDefaults()
                        .SetHatchPattern(HatchPatternType.PreDefined, HatchName)
                        .Associative = True
                        .PatternScale = HatchScale
                        .PatternAngle = HatchAngle
                        .Transparency = New Autodesk.AutoCAD.Colors.Transparency(CByte(Transparency))
                        .AppendLoop(HatchLoopTypes.Outermost, ObjectIdCollection)
                        .EvaluateHatch(True)
                    End With
                    Transaction.Commit()
                    Transaction.Dispose()
                End Using
                Return Hatch
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Adiciona hachura na entidade
        ''' </summary>
        ''' <param name="ObjectID">ID da entidade</param>
        ''' <param name="HatchName">Nome da hachura</param>
        ''' <param name="HatchScale">Escala</param>
        ''' <param name="Transparency">Transparência</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AddHatch(ByVal ObjectID As ObjectId, ByVal HatchName As String, ByVal HatchScale As Double, HatchAngle As Double, Optional Transparency As Integer = 0) As Hatch
            Try
                Transparency = Math.Truncate(255 * (100 - Transparency) / 100)
                Dim Hatch As Hatch = Nothing
                Dim acadDocument As Document = Application.DocumentManager.MdiActiveDocument
                Dim acadDatabase As Database = acadDocument.Database
                Using Transaction As Transaction = acadDatabase.TransactionManager.StartTransaction()
                    Dim BlockTable As BlockTable = Transaction.GetObject(acadDatabase.BlockTableId, OpenMode.ForRead)
                    Dim BlockTableRecord As BlockTableRecord = Transaction.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim ObjectIdCollection As ObjectIdCollection = New ObjectIdCollection()
                    ObjectIdCollection.Add(ObjectID)
                    Hatch = New Hatch()
                    BlockTableRecord.AppendEntity(Hatch)
                    Transaction.AddNewlyCreatedDBObject(Hatch, True)
                    With Hatch
                        .SetDatabaseDefaults()
                        .SetHatchPattern(HatchPatternType.PreDefined, HatchName)
                        .Associative = True
                        .PatternScale = HatchScale
                        .PatternAngle = HatchAngle
                        .Transparency = New Autodesk.AutoCAD.Colors.Transparency(CByte(Transparency))
                        .AppendLoop(HatchLoopTypes.Outermost, ObjectIdCollection)
                        .EvaluateHatch(True)
                    End With
                    Transaction.Commit()
                    Transaction.Dispose()
                End Using
                Return Hatch
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Adiciona hachura na entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="HatchName">Nome da hachura</param>
        ''' <param name="HatchScale">Escala</param>
        ''' <param name="Transparency">Transparência</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AddHatch(ByVal Entity As Entity, ByVal HatchName As String, ByVal HatchScale As Double, HatchAngle As Double, Optional Transparency As Integer = 0) As Hatch
            Try
                Transparency = Math.Truncate(255 * (100 - Transparency) / 100)
                Dim Hatch As Hatch = Nothing
                Dim acadDocument As Document = Application.DocumentManager.MdiActiveDocument
                Dim acadDatabase As Database = acadDocument.Database
                Using Transaction As Transaction = acadDatabase.TransactionManager.StartTransaction()
                    Dim BlockTable As BlockTable = Transaction.GetObject(acadDatabase.BlockTableId, OpenMode.ForRead)
                    Dim BlockTableRecord As BlockTableRecord = Transaction.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim ObjectIdCollection As ObjectIdCollection = New ObjectIdCollection()
                    ObjectIdCollection.Add(Entity.ObjectId)
                    Hatch = New Hatch()
                    BlockTableRecord.AppendEntity(Hatch)
                    Transaction.AddNewlyCreatedDBObject(Hatch, True)
                    With Hatch
                        .SetDatabaseDefaults()
                        .SetHatchPattern(HatchPatternType.PreDefined, HatchName)
                        .Associative = True
                        .PatternScale = HatchScale
                        .PatternAngle = HatchAngle
                        .Transparency = New Autodesk.AutoCAD.Colors.Transparency(CByte(Transparency))
                        .AppendLoop(HatchLoopTypes.Outermost, ObjectIdCollection)
                        .EvaluateHatch(True)
                    End With
                    Transaction.Commit()
                    Transaction.Dispose()
                End Using
                Return Hatch
            Catch
                Return Nothing
            End Try
        End Function

    End Class

End Namespace