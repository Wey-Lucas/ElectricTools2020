'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.Text
Imports Autodesk.AutoCAD.Runtime
Imports System.Collections.Generic

Namespace Engine2

    ''' <summary>
    ''' Cria uma janela de visualização em um bloco ou XRef (Similar ao comando XClip)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class XClip

        ''' <summary>
        ''' Cria a janela de visualização no bloco
        ''' </summary>
        ''' <param name="BlockReference">Bloco</param>
        ''' <param name="P1">Ponto inicial da janela</param>
        ''' <param name="P2">Ponto oposto da janela</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function XClip(BlockReference As BlockReference, P1 As Point3d, P2 As Point3d) As Boolean
            Dim SB As New StringBuilder
            With SB
                .Append("_.XCLIP")
                .Append(vbLf)
                .Append("(Handent " & Chr(34) & BlockReference.Handle.ToString & Chr(34) & ")")
                .Append(vbLf)
                .Append(vbLf)
                .Append("_NEW")
                .Append(vbLf)
                .Append("_RECTANGULAR")
                .Append(vbLf)
                .Append("(List " & P1.X & " " & P1.Y & " " & P1.Z & ")")
                .Append(vbLf)
                .Append("(List " & P2.X & " " & P2.Y & " " & P2.Z & ")")
                .Append(vbLf)
                .Append("(Princ)")
                .Append(vbLf)
            End With
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute(SB.ToString, True, False, False)
            Return True
        End Function

        ''' <summary>
        ''' Limites de um XClip
        ''' </summary>
        ''' <param name="BlockReference"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function XClipLimits(BlockReference As BlockReference) As Point2dCollection
            XClipLimits = Nothing
            Const FilterDictName As String = "ACAD_FILTER"
            Const SpatialName As String = "SPATIAL"
            Dim Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor = Document.Editor
            Using Transaction = Document.TransactionManager.StartTransaction()
                BlockReference = TryCast(Transaction.GetObject(BlockReference.ObjectId, OpenMode.ForRead), BlockReference)
                Dim found As Boolean = False
                If IsNothing(BlockReference) = False AndAlso BlockReference.ExtensionDictionary.IsNull = False Then
                    Dim ExtensionDictionary = TryCast(Transaction.GetObject(BlockReference.ExtensionDictionary, OpenMode.ForRead), DBDictionary)
                    If ExtensionDictionary IsNot Nothing AndAlso ExtensionDictionary.Contains(FilterDictName) Then
                        Dim DBDictionary = TryCast(Transaction.GetObject(ExtensionDictionary.GetAt(FilterDictName), OpenMode.ForRead), DBDictionary)
                        If DBDictionary IsNot Nothing Then
                            If DBDictionary.Contains(SpatialName) Then
                                Dim fil = TryCast(Transaction.GetObject(DBDictionary.GetAt(SpatialName), OpenMode.ForRead), SpatialFilter)
                                If fil IsNot Nothing Then
                                    'Dim ext = fil.GetQueryBounds()
                                    XClipLimits = fil.Definition.GetPoints()
                                End If
                            End If
                        End If
                    End If
                End If
                Transaction.Commit()
            End Using
            Return XClipLimits
        End Function



        'APRESENTA ERRO APÓS A CRIAÇÃO DO XCLIP \ NÃO PERMITE SOBREPOSIÇÃO DE PONTOS
        ' ''' <summary>
        ' ''' Cria a janela de visualização no bloco
        ' ''' </summary>
        ' ''' <param name="BlockReference">Bloco</param>
        ' ''' <param name="P1">Ponto inicial da janela</param>
        ' ''' <param name="P2">Ponto oposto da janela</param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function XClip(BlockReference As BlockReference, P1 As Point3d, P2 As Point3d) As Boolean
        '    Try
        '        Const DictName As String = "ACAD_FILTER"
        '        Const SpatialName As String = "SPATIAL"
        '        Dim Document As Document = Application.DocumentManager.MdiActiveDocument
        '        Dim Database As Database = Document.Database
        '        Dim Editor As Editor = Document.Editor
        '        Dim SpatialFilterDefinition As Filters.SpatialFilterDefinition
        '        Dim DBDictionary As DBDictionary
        '        Dim Points As New List(Of Point3d)
        '        Dim Point2dCollection As Point2dCollection
        '        Points.Add(P1)
        '        Points.Add(P2)
        '        P1 = Engine.MyGeometry.MinPoint3d(Points)
        '        P2 = Engine.MyGeometry.MaxPoint3d(Points)
        '        Using Editor.Document.LockDocument
        '            Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
        '                BlockReference = TryCast(Transaction.GetObject(BlockReference.ObjectId, OpenMode.ForWrite), BlockReference)
        '                Point2dCollection = New Point2dCollection
        '                Point2dCollection.Add(Engine.MyGeometry.Point3dToPoint2d(P1))
        '                Point2dCollection.Add(Engine.MyGeometry.Point3dToPoint2d(P2))
        '                SpatialFilterDefinition = New Filters.SpatialFilterDefinition(Point2dCollection, Vector3d.ZAxis, 0, 0, 0, True)
        '                Using SpatialFilter As New Filters.SpatialFilter
        '                    SpatialFilter.Definition = SpatialFilterDefinition
        '                    If BlockReference.ExtensionDictionary.IsNull = True Then
        '                        BlockReference.UpgradeOpen()
        '                        BlockReference.CreateExtensionDictionary()
        '                        BlockReference.DowngradeOpen()
        '                    End If
        '                    DBDictionary = Transaction.GetObject(BlockReference.ExtensionDictionary, OpenMode.ForWrite)
        '                    If DBDictionary.Contains(DictName) = True Then
        '                        Using FilterDBDictionary As DBDictionary = Transaction.GetObject(DBDictionary.GetAt(DictName), OpenMode.ForWrite)
        '                            If FilterDBDictionary.Contains(SpatialName) = True Then
        '                                FilterDBDictionary.Remove(SpatialName)
        '                            End If
        '                            FilterDBDictionary.SetAt(SpatialName, SpatialFilter)
        '                        End Using
        '                    Else
        '                        Using FilterDBDictionary As New DBDictionary
        '                            DBDictionary.SetAt(DictName, FilterDBDictionary)
        '                            Transaction.AddNewlyCreatedDBObject(FilterDBDictionary, True)
        '                            FilterDBDictionary.SetAt(SpatialName, SpatialFilter)
        '                        End Using
        '                    End If
        '                    Transaction.AddNewlyCreatedDBObject(SpatialFilter, True)
        '                End Using
        '                Transaction.Commit()
        '            End Using
        '        End Using
        '        Editor.Regen()
        '        Return True
        '    Catch
        '        Return False
        '    End Try
        'End Function

    End Class

End Namespace





