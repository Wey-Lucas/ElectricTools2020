Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.Drawing
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Runtime.Interop
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports ErrorStatus = Autodesk.AutoCAD.Runtime.ErrorStatus
Imports acApp = Autodesk.AutoCAD.ApplicationServices.Application
Imports Autodesk.AutoCAD.ApplicationServices.DocumentExtension
Imports ElectricTools2020.Engine2.AcadInterface
Imports System.Runtime.InteropServices

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Classe com recursos de interação avançada com entidades
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EntityInteration

        ''' <summary>
        ''' Ordem da entidade
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eDrawOrder
            ''' <summary>
            ''' Em cima
            ''' </summary>
            ''' <remarks></remarks>
            Top = 0

            ''' <summary>
            ''' Em baixo
            ''' </summary>
            ''' <remarks></remarks>
            Botton = 1
        End Enum

        'AUTOCAD 2014/2018
        <System.Security.SuppressUnmanagedCodeSecurity>
        <DllImport("accore.dll", CallingConvention:=CallingConvention.Cdecl)>
        Private Shared Function acedRedraw(ByVal name As Long(), ByVal mode As Int32) As Int32
        End Function
        ''' <summary>
        ''' Redesenha a tela
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Redraw2()
            acedRedraw(Nothing, 1)
        End Sub

        ''' <summary>
        ''' Redesenha a tela (Command)
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Redraw()
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            If Document.CommandInProgress = "" Then
                Document.SendStringToExecute("(command " & Chr(34) & "_.REDRAW" & Chr(34) & " )(princ) ", True, False, False)
            End If
        End Sub

        ''' <summary>
        ''' Retorna se um ponto esta contido em uma entidade
        ''' </summary>
        ''' <param name="Entity">Entidade a ser avaliada</param>
        ''' <param name="Point3d">Ponto procurado</param>
        ''' <param name="AllowedClassCollection">Filtro de entidades</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsInsidePoint(Entity As Entity, Point3d As Point3d, Optional AllowedClassCollection As List(Of System.Type) = Nothing) As Boolean
            If IsNothing(AllowedClassCollection) = False Then
                If AllowedClassCollection.Contains(Entity.GetType) = False Then
                    Return False
                End If
            End If
            Dim Curve As Curve
            Dim CalcPoint As Point3d
            Dim Curve3d As Curve3d = Nothing
            Dim Polyline As Polyline
            Dim DBObjectCollection As New DBObjectCollection
            Dim DBPoint As DBPoint
            Dim DBText As DBText
            Dim MText As MText
            Select Case Entity.GetType.Name
                Case "BlockReference", "Hatch"
                    Try
                        Entity.Explode(DBObjectCollection)
                        For Each DBObject As DBObject In DBObjectCollection
                            Select Case DBObject.GetType.Name
                                Case "BlockReference", "Hatch"
                                    Entity = DBObject
                                    If IsInsidePoint(Entity, Point3d) = True Then
                                        Return True
                                    End If
                                Case Else
                                    Entity = DBObject
                                    If IsInsidePoint(Entity, Point3d) = True Then
                                        Return True
                                    End If
                            End Select
                        Next
                    Catch
                        Return False
                    End Try
                Case "Polyline"
                    Polyline = Entity
                    For Index As Integer = 0 To Polyline.NumberOfVertices - 1
                        Select Case Polyline.GetSegmentType(Index)
                            Case SegmentType.Arc
                                Curve3d = Polyline.GetArcSegmentAt(Index)
                            Case SegmentType.Line
                                Curve3d = Polyline.GetLineSegmentAt(Index)
                        End Select
                        If IsNothing(Curve3d) = False Then
                            If Curve3d.IsOn(Point3d) = True Then
                                Return True
                            End If
                        End If
                    Next
                Case "Line", "Spline", "Polyline2d", "Polyline3d", "Circle", "Arc", "Ellipse"
                    Curve = Entity
                    CalcPoint = Curve.GetClosestPointTo(Point3d, False)
                    If Point3d.DistanceTo(CalcPoint) < Tolerance.Global.EqualPoint Then
                        Return True
                    End If
                Case "DBPoint"
                    DBPoint = Entity
                    If Point3d.DistanceTo(DBPoint.Position) < Tolerance.Global.EqualPoint Then
                        Return True
                    End If
                Case "DBText"
                    DBText = Entity
                    If Point3d.DistanceTo(DBText.Position) < Tolerance.Global.EqualPoint Then
                        Return True
                    End If
                Case "MText"
                    MText = Entity
                    If Point3d.DistanceTo(MText.Location) < Tolerance.Global.EqualPoint Then
                        Return True
                    End If
                Case Else
                    Throw New System.Exception("O tipo '" & Entity.GetType.Name & "' não é válido.")
            End Select
            Return False
        End Function

        ' ''' <summary>
        ' ''' Determina se um ponto esta sobre uma polyline
        ' ''' </summary>
        ' ''' <param name="Polyline">Polyline</param>
        ' ''' <param name="Point3d">Ponto</param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function IsInsidePoint(Polyline As Polyline, Point3d As Point3d) As Boolean
        '    Try
        '        Dim Curve3d As Curve3d = Nothing
        '        For Index As Integer = 0 To Polyline.NumberOfVertices - 1
        '            Select Case Polyline.GetSegmentType(Index)
        '                Case SegmentType.Arc
        '                    Curve3d = Polyline.GetArcSegmentAt(Index)
        '                Case SegmentType.Line
        '                    Curve3d = Polyline.GetLineSegmentAt(Index)
        '            End Select
        '            If IsNothing(Curve3d) = False Then
        '                If Curve3d.IsOn(Point3d) = True Then
        '                    Return True
        '                End If
        '            End If
        '        Next
        '        Return False
        '    Catch
        '        Return False
        '    End Try
        'End Function

        ' ''' <summary>
        ' ''' Complemento da função SelectAtPoint
        ' ''' </summary>
        ' ''' <param name="BlockReference">Bloco</param>
        ' ''' <param name="Point3d">Ponto</param>
        ' ''' <param name="AllowedClassCollection">Classes com os tipos de entidades a serem observadas</param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Shared Function IsInsidePoint(BlockReference As BlockReference, Point3d As Point3d, AllowedClassCollection As List(Of System.Type)) As Boolean
        '    Try
        '        Dim DBObjectCollection As New DBObjectCollection
        '        Dim Curve As Curve
        '        Dim CalcPoint As Point3d
        '        Dim DBPoint As DBPoint
        '        BlockReference.Explode(DBObjectCollection)
        '        For Each DBObject As DBObject In DBObjectCollection
        '            Select Case DBObject.GetType
        '                Case GetType(Line), GetType(Spline), GetType(Polyline), GetType(Polyline2d), GetType(Polyline3d), GetType(Circle), GetType(Arc), GetType(Ellipse)
        '                    Curve = DBObject
        '                    CalcPoint = Curve.GetClosestPointTo(Point3d, False)
        '                    If Point3d.DistanceTo(CalcPoint) < Tolerance.Global.EqualPoint Then
        '                        Return True
        '                    End If
        '                Case GetType(BlockReference)
        '                    If IsInsidePoint(DBObject, Point3d, AllowedClassCollection) = True Then
        '                        Return True
        '                    End If
        '                Case GetType(DBPoint)
        '                    DBPoint = DBObject
        '                    If Point3d.DistanceTo(DBPoint.Position) < Tolerance.Global.EqualPoint Then
        '                        Return True
        '                    End If
        '                Case Else
        '                    Throw New System.Exception("O tipo '" & DBObject.GetType.Name & "' não é válido.")
        '            End Select
        '        Next
        '        Return False
        '    Catch
        '        Return False
        '    End Try
        'End Function

        ''' <summary>
        ''' Determina se um ponto esta contido dentro de uma entidade
        ''' </summary>
        ''' <param name="Entity"></param>
        ''' <param name="Point3d"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsInsideInternalPoint(Entity As Entity, Point3d As Point3d) As Boolean
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim Editor As Editor = Document.Editor
                Dim DBObjectCollection As DBObjectCollection
                DBObjectCollection = Editor.TraceBoundary(Point3d, True)
                If DBObjectCollection.Count > 0 Then
                    Return DBObjectCollection.Contains(Entity)
                Else
                    Return False
                End If
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Compara curvas
        ''' </summary>
        ''' <param name="Curve1">Curva</param>
        ''' <param name="curve2">Curva para comparação</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CurveIsEquals(Curve1 As Curve, Curve2 As Curve) As Boolean
            Return Curve1.GetGeCurve.IsEqualTo(Curve2.GetGeCurve)
        End Function

        ''' <summary>
        ''' Calcula os pontos do retângulo que engloba as entidades
        ''' </summary>
        ''' <param name="DBObjectCollection">Coleção de entidades a serem avaliadas</param>
        ''' <returns>Rectangle3d</returns>
        ''' <remarks></remarks>
        Public Shared Function GetRectangle3d(DBObjectCollection As DBObjectCollection) As Rectangle3d
            Dim Points As New Point3dCollection
            Dim Extents3d As Extents3d
            Dim LowerRight As Point3d
            Dim UpperLefty As Point3d
            Dim LowerLefty As Point3d
            Dim UpperRight As Point3d
            Dim Z As Double
            For Each Entity As Entity In DBObjectCollection
                Extents3d = Entity.GeometricExtents
                If Points.Contains(Extents3d.MinPoint) = False Then
                    Points.Add(Extents3d.MinPoint)
                End If
                If Points.Contains(Extents3d.MaxPoint) = False Then
                    Points.Add(Extents3d.MaxPoint)
                End If
            Next
            Z = 0
            UpperRight = Engine2.Geometry.MaxPoint3d(Points)
            UpperRight = New Point3d(UpperRight.X, UpperRight.Y, Z)
            LowerLefty = Engine2.Geometry.MinPoint3d(Points)
            LowerLefty = New Point3d(LowerLefty.X, LowerLefty.Y, Z)
            UpperLefty = New Point3d(LowerLefty.X, UpperRight.Y, Z)
            LowerRight = New Point3d(UpperRight.X, LowerLefty.Y, Z)
            Return New Rectangle3d(UpperLefty, UpperRight, LowerLefty, LowerRight)
        End Function

        ''' <summary>
        ''' Desfaz a seleção de entidades na tela
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub Unselect()
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim SelectionSet As Autodesk.AutoCAD.EditorInput.SelectionSet = Nothing
            Editor.SetImpliedSelection(SelectionSet)
        End Sub

        ''' <summary>
        ''' Cria a seleção de entidades na tela
        ''' </summary>
        ''' <param name="ObjectIdCollection">Coleção de ID´s</param>
        ''' <remarks></remarks>
        Public Shared Sub [Select](ObjectIdCollection As ObjectIdCollection)
            If ObjectIdCollection.Count > 0 Then
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim Editor As Editor = Document.Editor
                Dim ObjectIds(ObjectIdCollection.Count - 1) As ObjectId
                For Index As Integer = 0 To ObjectIdCollection.Count - 1
                    ObjectIds(Index) = ObjectIdCollection.Item(Index)
                Next
                Editor.SetImpliedSelection(ObjectIds)
            Else
                Engine2.EntityInteration.Unselect()
            End If
        End Sub

        ''' <summary>
        ''' Cria a seleção de entidades na tela
        ''' </summary>
        ''' <param name="DbObjectCollection">Coleção de entidades</param>
        ''' <remarks></remarks>
        Public Shared Sub [Select](DbObjectCollection As DBObjectCollection)
            If DbObjectCollection.Count > 0 Then
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim Editor As Editor = Document.Editor
                Dim ObjectIds(DbObjectCollection.Count - 1) As ObjectId
                For Index As Integer = 0 To DbObjectCollection.Count - 1
                    ObjectIds(Index) = DbObjectCollection.Item(Index).ObjectId
                Next
                Editor.SetImpliedSelection(ObjectIds)
            Else
                Engine2.EntityInteration.Unselect()
            End If
        End Sub

        ''' <summary>
        ''' Obtem o centro de uma polyline
        ''' </summary>
        ''' <param name="Polyline">Polyline</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetCentroid(ByVal Polyline As Polyline) As Object
            Try
                Dim ReferencePoint As Point2d
                Dim RunningX As Double
                Dim RunningY As Double
                Dim second_factor As Double
                Dim Point1 As Point2d
                Dim Point2 As Point2d
                Dim X As Double
                Dim Y As Double
                For Counter As Integer = 0 To Polyline.NumberOfVertices - 2
                    Point1 = Polyline.GetPoint2dAt(Counter)
                    Point2 = Polyline.GetPoint2dAt(Counter + 1)
                    If Counter = 0 Then ReferencePoint = Point1
                    second_factor = (Point1.X - ReferencePoint.X) * (Point2.Y - ReferencePoint.Y) - (Point2.X - ReferencePoint.X) * (Point1.Y - ReferencePoint.Y)
                    RunningX += ((Point1.X - ReferencePoint.X) + (Point2.X - ReferencePoint.X)) * second_factor
                    RunningY += ((Point1.Y - ReferencePoint.Y) + (Point2.Y - ReferencePoint.Y)) * second_factor
                Next
                X = RunningX / (6 * Polyline.Area) + ReferencePoint.X
                Y = RunningY / (6 * Polyline.Area) + ReferencePoint.Y
                Return New Point3d(X, Y, Polyline.Elevation)
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Determina se um objeto de dados é entidade
        ''' </summary>
        ''' <param name="DbObject"></param>
        ''' <param name="GraphicalOnly">Determina se apenas entidades visíveis serão consideradas</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsEntity(DbObject As DBObject, Optional GraphicalOnly As Boolean = False) As Boolean
            If GraphicalOnly = False Then
                If IsNothing(Engine2.ConvertObject.DBObjectToEntity(DbObject)) = False Then
                    Return True
                Else
                    Return False
                End If
            Else
                If IsNothing(Engine2.ConvertObject.DBObjectToEntity(DbObject)) = False And DbObject.GetType.Name.Equals("Wipeout") = False Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Function

        ''' <summary>
        ''' Ordem de desenho
        ''' </summary>
        ''' <param name="Reference">ID da entidade a ser ordenada</param>
        ''' <param name="DrawOrder">Posição</param>
        ''' <param name="Target">ID da entidade de referência</param>
        ''' <remarks></remarks>
        Public Shared Sub DrawOrder(Reference As ObjectId, DrawOrder As eDrawOrder, Optional Target As ObjectId = Nothing)
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTableRecord As BlockTableRecord
            Dim DrawOrderTable As DrawOrderTable
            Dim Collection As New ObjectIdCollection()
            Dim Entity As Entity
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    Entity = Transaction.GetObject(Reference, OpenMode.ForRead)
                    BlockTableRecord = Transaction.GetObject(Entity.BlockId, OpenMode.ForRead)
                    DrawOrderTable = Transaction.GetObject(BlockTableRecord.DrawOrderTableId, OpenMode.ForWrite)
                    Collection.Add(Entity.ObjectId)
                    Select Case DrawOrder
                        Case eDrawOrder.Top
                            If Target.IsValid = False Then
                                DrawOrderTable.MoveToTop(Collection)
                            Else
                                DrawOrderTable.MoveAbove(Collection, Target)
                            End If
                        Case eDrawOrder.Botton
                            If Target.IsValid = False Then
                                DrawOrderTable.MoveToBottom(Collection)
                            Else
                                DrawOrderTable.MoveBelow(Collection, Target)
                            End If
                    End Select
                    Transaction.Commit()
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Ordem de desenho
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Reference">ID da entidade a ser ordenada</param>
        ''' <param name="DrawOrder">Posição</param>
        ''' <param name="Target">ID da entidade de referência</param>
        ''' <remarks></remarks>
        Public Shared Sub DrawOrder(Transaction As Transaction, Reference As ObjectId, DrawOrder As eDrawOrder, Optional Target As ObjectId = Nothing)
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTableRecord As BlockTableRecord
            Dim DrawOrderTable As DrawOrderTable
            Dim Collection As New ObjectIdCollection()
            Dim Entity As Entity
            Entity = Transaction.GetObject(Reference, OpenMode.ForRead)
            BlockTableRecord = Transaction.GetObject(Entity.BlockId, OpenMode.ForRead)
            DrawOrderTable = Transaction.GetObject(BlockTableRecord.DrawOrderTableId, OpenMode.ForWrite)
            Collection.Add(Entity.ObjectId)
            Select Case DrawOrder
                Case eDrawOrder.Top
                    If Target.IsValid = False Then
                        DrawOrderTable.MoveToTop(Collection)
                    Else
                        DrawOrderTable.MoveAbove(Collection, Target)
                    End If
                Case eDrawOrder.Botton
                    If Target.IsValid = False Then
                        DrawOrderTable.MoveToBottom(Collection)
                    Else
                        DrawOrderTable.MoveBelow(Collection, Target)
                    End If
            End Select
        End Sub

        ''' <summary>
        ''' Ordem de desenho
        ''' </summary>
        ''' <param name="Collection">Coleção a ser ordenada</param>
        ''' <param name="DrawOrder">Posição</param>
        ''' <param name="Target">ID da entidade de referência</param>
        ''' <remarks></remarks>
        Public Shared Sub DrawOrder(Collection As ObjectIdCollection, DrawOrder As eDrawOrder, Optional Target As ObjectId = Nothing)
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTableRecord As BlockTableRecord
            Dim DrawOrderTable As DrawOrderTable
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForRead)
                    DrawOrderTable = Transaction.GetObject(BlockTableRecord.DrawOrderTableId, OpenMode.ForWrite)
                    Select Case DrawOrder
                        Case eDrawOrder.Top
                            If Target.IsValid = False Then
                                DrawOrderTable.MoveToTop(Collection)
                            Else
                                DrawOrderTable.MoveAbove(Collection, Target)
                            End If
                        Case eDrawOrder.Botton
                            If Target.IsValid = False Then
                                DrawOrderTable.MoveToBottom(Collection)
                            Else
                                DrawOrderTable.MoveBelow(Collection, Target)
                            End If
                    End Select
                    Transaction.Commit()
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Ordem de desenho
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Collection">Coleção a ser ordenada</param>
        ''' <param name="DrawOrder">Posição</param>
        ''' <param name="Target">ID da entidade de referência</param>
        ''' <remarks></remarks>
        Public Shared Sub DrawOrder(Transaction As Transaction, Collection As ObjectIdCollection, DrawOrder As eDrawOrder, Optional Target As ObjectId = Nothing)
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTableRecord As BlockTableRecord
            Dim DrawOrderTable As DrawOrderTable
            BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForRead)
            DrawOrderTable = Transaction.GetObject(BlockTableRecord.DrawOrderTableId, OpenMode.ForWrite)
            Select Case DrawOrder
                Case eDrawOrder.Top
                    If Target.IsValid = False Then
                        DrawOrderTable.MoveToTop(Collection)
                    Else
                        DrawOrderTable.MoveAbove(Collection, Target)
                    End If
                Case eDrawOrder.Botton
                    If Target.IsValid = False Then
                        DrawOrderTable.MoveToBottom(Collection)
                    Else
                        DrawOrderTable.MoveBelow(Collection, Target)
                    End If
            End Select
        End Sub

        ''' <summary>
        '''  Converte LineSegment2d em Line
        ''' </summary>
        ''' <param name="LineSegment2d">LineSegment2d</param>
        ''' <returns>Line</returns>
        ''' <remarks></remarks>
        Public Shared Function LineSegment2dToLine(LineSegment2d As LineSegment2d) As Line
            Try
                Dim Line As Line
                Dim Interval As Interval = LineSegment2d.GetInterval
                Dim StartParam As Double = Interval.LowerBound
                Dim EndParam As Double = Interval.UpperBound
                Dim StarPoint2d As Point2d = LineSegment2d.EvaluatePoint(StartParam)
                Dim EndPoint2d As Point2d = LineSegment2d.EvaluatePoint(EndParam)
                Dim StartPoint As Point3d = New Point3d(StarPoint2d.X, StarPoint2d.Y, 0)
                Dim EndPoint As Point3d = New Point3d(EndPoint2d.X, EndPoint2d.Y, 0)
                Line = New Line(StartPoint, EndPoint)
                Return Line
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Converte ArcSegment2d em Arc
        ''' </summary>
        ''' <param name="CircularArc2d">CircularArc2d</param>
        ''' <returns>Arc</returns>
        ''' <remarks></remarks>
        Public Shared Function CircularArc2dToArc(CircularArc2d As CircularArc2d) As Arc
            Try
                Dim Arc As Arc
                Dim Interval As Interval = CircularArc2d.GetInterval
                Dim StartParam As Double = Interval.LowerBound
                Dim EndParam As Double = Interval.UpperBound
                Dim StartPoint As Point2d = CircularArc2d.EvaluatePoint(StartParam)
                Dim EndPoint As Point2d = CircularArc2d.EvaluatePoint(EndParam)
                Dim Point2d As Point2d = CircularArc2d.Center
                Dim StartAngle As Double = (New Line2d(Point2d, StartPoint)).Direction.Angle
                Dim EndAngle As Double = (New Line2d(Point2d, EndPoint)).Direction.Angle
                Dim Center As New Point3d(Point2d.X, Point2d.Y, 0)
                If CircularArc2d.IsClockWise = True Then
                    Arc = New Arc(Center, CircularArc2d.Radius, EndAngle, StartAngle)
                Else
                    Arc = New Arc(Center, CircularArc2d.Radius, StartAngle, EndAngle)
                End If
                Return Arc
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retorna o comprimento de uma entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <returns>SizeF</returns>
        ''' <remarks></remarks>
        Public Shared Function Size(Entity As Entity) As SizeF
            Size = New SizeF(0, 0)
            Try
                Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
                Dim Database As Database = Editor.Document.Database
                Dim Width As Double = 0
                Dim Height As Double = 0
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        Dim Extents3d As Extents3d = Entity.GeometricExtents
                        Width = New Point2d(Extents3d.MinPoint.X, 0).GetDistanceTo(New Point2d(Extents3d.MaxPoint.X, 0))
                        Height = New Point2d(0, Extents3d.MaxPoint.Y).GetDistanceTo(New Point2d(0, Extents3d.MinPoint.Y))
                        Size = New SizeF(Width, Height)
                        Transaction.Abort()
                    End Using
                End Using
            Catch
                Size = Nothing
            End Try
            Return Size
        End Function

        '''' <summary>
        '''' Retorna o centro da entidade
        '''' </summary>
        '''' <param name="Entity">Entidade</param>
        '''' <returns>SizeF</returns>
        '''' <remarks></remarks>
        'Public Shared Function Center(Entity As Entity) As Object
        '    Center = Nothing
        '    Try
        '        Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        '        Dim Database As Database = Editor.Document.Database
        '        Dim Width As Double = 0
        '        Dim Height As Double = 0
        '        Using Editor.Document.LockDocument
        '            Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
        '                Dim Extents3d As Extents3d = Entity.GeometricExtents
        '                Center = Engine2.Geometry.MidPoint(Extents3d.MinPoint, Extents3d.MaxPoint)
        '                Transaction.Abort()
        '            End Using
        '        End Using
        '    Catch
        '        Center = Nothing
        '    End Try
        '    Return Center
        'End Function

        ''' <summary>
        ''' Destaca a entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function HighlightOn(Entity As Entity) As Boolean
            Try
                Using Database As Database = HostApplicationServices.WorkingDatabase()
                    Using Transaction As Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction()
                        Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                        Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                            Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
                            Entity.Highlight()
                        End Using
                    End Using
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Destaca a entidade
        ''' </summary>
        ''' <param name="ObjectId">ID</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function HighlightOn(ObjectId As ObjectId) As Boolean
            Try
                Dim Entity As Entity
                Using Database As Database = HostApplicationServices.WorkingDatabase()
                    Using Transaction As Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction()
                        Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                        Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                            Entity = Transaction.GetObject(ObjectId, OpenMode.ForRead)
                            Entity.Highlight()
                        End Using
                    End Using
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Remove o destaque da entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function HighlightOff(Entity As Entity) As Boolean
            Try
                Using Database As Database = HostApplicationServices.WorkingDatabase()
                    Using Transaction As Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction()
                        Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                        Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                            Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
                            Entity.Unhighlight()
                        End Using
                    End Using
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Remove o destaque da entidade
        ''' </summary>
        ''' <param name="ObjectId">Id</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function HighlightOff(ObjectId As ObjectId) As Boolean
            Try
                Dim Entity As Entity
                Using Database As Database = HostApplicationServices.WorkingDatabase()
                    Using Transaction As Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction()
                        Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                        Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                            Entity = Transaction.GetObject(ObjectId, OpenMode.ForRead)
                            Entity.Unhighlight()
                        End Using
                    End Using
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Determina se um ponto faz parte da entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Point3d">Ponto a ser avaliado</param>
        ''' <param name="Extend">Extender</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PointAtEntity(Entity As Entity, Point3d As Point3d, Optional Extend As Boolean = True) As Boolean
            Try
                Dim Curve As Curve = Entity
                If ((Curve.GetClosestPointTo(Point3d, Extend) - Point3d).Length <= Tolerance.Global.EqualPoint) Then
                    Return True
                Else
                    Return False
                End If
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Determina se um ponto faz parte da entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="Point3dCollection">Ponto a ser avaliado</param>
        ''' <param name="Extend">Extender</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PointAtEntity(Entity As Entity, Point3dCollection As Point3dCollection, Optional Extend As Boolean = True) As Boolean
            Try
                Dim Curve As Curve = Entity
                For Each Point3d As Point3d In Point3dCollection
                    If (Curve.GetClosestPointTo(Point3d, Extend) - Point3d).Length > Tolerance.Global.EqualPoint Then
                        Return False
                    End If
                Next
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Obtem as sub-entidades de forma mais pura (onde não é possível mais explosão). 
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="SubItems">Determina se os sub itens serão explodidos en todos os níveis</param>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <param name="FilterEntity">Determina as entidades válidas para avaliação</param>
        ''' <param name="EntitysException">O tipo de entidade especificada será desconsiderada na explosão</param>
        ''' <param name="Tolerance">Tolerância máxima para considerar a entidade válida para o ponto de referência (Tolerance.Global.EqualPoint)</param>
        ''' <param name="UseExactPointReference">Determina se será retornado apenas a entidade onde o ponto de referência se encontre ou se será retornado apenas a entidade mais próxima ao ponto de referência</param>
        ''' <returns>DBObjectCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function GetSubEntitys(Entity As Entity, Optional SubItems As Boolean = False, Optional PointReference As Object = Nothing, Optional FilterEntity As ArrayList = Nothing, Optional EntitysException As ArrayList = Nothing, Optional Tolerance As Object = Nothing, Optional UseExactPointReference As Boolean = False) As DBObjectCollection
            GetSubEntitys = New DBObjectCollection
            Try
                Dim Database As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
                Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
                Dim TempDBObjectCollection As New DBObjectCollection
                Dim DBObjectCollection As New DBObjectCollection
                Dim FilterDBObjectCollection As New DBObjectCollection
                Dim PerpendicularInforCollection As List(Of PerpendicularInfor)
                Dim PerpendicularInfor As PerpendicularInfor = Nothing
                If Engine2.EntityInteration.IsExplosible(Entity) = True Then
                    Entity.Explode(TempDBObjectCollection)
                    If SubItems = True Then
                        While TempDBObjectCollection.Count <> 0
                            Entity = TempDBObjectCollection.Item(0)
                            If Engine2.EntityInteration.IsExplosible(Entity, EntitysException) = True Then
                                FilterDBObjectCollection.Clear()
                                Entity.Explode(FilterDBObjectCollection)
                                For Each DBObject As DBObject In FilterDBObjectCollection
                                    Entity = DBObject
                                    If Engine2.EntityInteration.IsExplosible(Entity, EntitysException) = True Then
                                        TempDBObjectCollection.Add(Entity)
                                    Else
                                        If IsNothing(FilterEntity) = False Then
                                            If FilterEntity.Contains(Entity.GetType.Name) = True Then
                                                DBObjectCollection.Add(Entity)
                                            End If
                                        Else
                                            DBObjectCollection.Add(Entity)
                                        End If
                                    End If
                                Next
                            Else
                                If IsNothing(FilterEntity) = False Then
                                    If FilterEntity.Contains(Entity.GetType.Name) = True Then
                                        DBObjectCollection.Add(Entity)
                                    End If
                                Else
                                    DBObjectCollection.Add(Entity)
                                End If
                            End If
                            TempDBObjectCollection.RemoveAt(0)
                        End While
                    Else
                        If IsNothing(FilterEntity) = False Then
                            While TempDBObjectCollection.Count <> 0
                                Entity = TempDBObjectCollection.Item(0)
                                If FilterEntity.Contains(Entity.GetType.Name) = True Then
                                    DBObjectCollection.Add(Entity)
                                End If
                                TempDBObjectCollection.RemoveAt(0)
                            End While
                        Else
                            DBObjectCollection = TempDBObjectCollection
                        End If
                    End If
                Else
                    DBObjectCollection.Add(Entity)
                    Return DBObjectCollection
                End If
                If DBObjectCollection.Count = 0 Then
                    Return Nothing
                Else
                    If IsNothing(PointReference) = False Then
                        'Verifica se o ponto esta na entidade selecionada
                        FilterDBObjectCollection.Clear()
                        For Each DBObject As DBObject In DBObjectCollection
                            If Engine2.EntityInteration.PointAtEntity(Engine2.ConvertObject.DBObjectToEntity(DBObject), PointReference) = True Then
                                FilterDBObjectCollection.Add(DBObject)
                                Exit For
                            End If
                        Next
                        'Determina se a entidade próxima será considerada 
                        If UseExactPointReference = False Then
                            'Caso não esteja, calcula a que possui perpendicular mais proxima
                            If FilterDBObjectCollection.Count = 0 Then
                                PerpendicularInforCollection = New List(Of PerpendicularInfor)
                                For Each DBObject As DBObject In DBObjectCollection
                                    PerpendicularInfor = Engine2.Geometry.GetPerpendicular(DBObject, PointReference)
                                    If IsNothing(PerpendicularInfor) = False Then
                                        If IsNothing(Tolerance) = True Then
                                            PerpendicularInfor.Tag = PointReference
                                            PerpendicularInforCollection.Add(PerpendicularInfor)
                                        Else
                                            If PointReference.DistanceTo(PerpendicularInfor.PerpendicularPoint) <= Tolerance Then
                                                PerpendicularInfor.Tag = PointReference
                                                PerpendicularInforCollection.Add(PerpendicularInfor)
                                            End If
                                        End If
                                    End If
                                Next
                                If PerpendicularInforCollection.Count > 0 Then
                                    PerpendicularInforCollection.Sort(Function(PI1 As PerpendicularInfor, PI2 As PerpendicularInfor) PI1.PerpendicularPoint.DistanceTo(PI1.Tag).CompareTo(PI2.PerpendicularPoint.DistanceTo(PI2.Tag)))
                                    DBObjectCollection.Clear()
                                    DBObjectCollection.Add(PerpendicularInforCollection.Item(0).Curve)
                                End If
                            Else
                                DBObjectCollection = FilterDBObjectCollection
                            End If
                        End If
                    End If
                    GetSubEntitys = DBObjectCollection
                    Return GetSubEntitys
                End If
            Catch
                Return GetSubEntitys
            End Try
        End Function

        ''' <summary>
        ''' Retorna se uma entidade pode ser explodida
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="EntitysException">O tipo de entidade especificada será desconsiderada</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsExplosible(Entity As Entity, Optional EntitysException As ArrayList = Nothing) As Boolean
            Try
                If IsNothing(EntitysException) = True Then
                    Entity.Explode(New DBObjectCollection)
                    Return True
                Else
                    If EntitysException.Contains(Entity.GetType.Name) = False Then
                        Entity.Explode(New DBObjectCollection)
                        Return True
                    Else
                        Return False
                    End If
                End If
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Retorna se o segmento de uma polyline é arco
        ''' </summary>
        ''' <param name="Polyline">Polyline</param>
        ''' <param name="VerticeIndex">Index do vértice</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsArcSegment(Polyline As Polyline, VerticeIndex As Integer) As Boolean
            Select Case Polyline.GetSegmentType(VerticeIndex)
                Case SegmentType.Arc
                    Return True
            End Select
            Return False
        End Function

        ''' <summary>
        ''' Retorna a entidade resultante de um OffSet
        ''' </summary>
        ''' <param name="Curve">Entidade de origem</param>
        ''' <param name="OffSetValue">Valor do OffSet</param>
        ''' <returns></returns>
        ''' <remarks>DBObject\Nothing</remarks>
        Public Shared Function OffSet(Curve As Curve, OffSetValue As Double) As Object
            Try
                Return Curve.GetOffsetCurves(OffSetValue).Item(0)
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem a polyline com base em um arco
        ''' </summary>
        ''' <param name="Arc"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetPolylineFromArc(Arc As Arc) As Polyline
            Dim Polyline As Polyline
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim BlockTableRecord As BlockTableRecord
            Dim Point2d As Point2d
            Dim Point2dCollection As New Point2dCollection
            Dim Angle As Double
            Dim Bulge As Double
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    Try
                        Point2dCollection.Add(Arc.StartPoint.Convert2d(New Plane(Point3d.Origin, Vector3d.ZAxis)))
                        Point2dCollection.Add(Arc.EndPoint.Convert2d(New Plane(Point3d.Origin, Vector3d.ZAxis)))
                        Polyline = New Polyline(Point2dCollection.Count)
                        Polyline.Layer = Arc.Layer
                        Angle = (Arc.TotalAngle / 2.0)
                        Bulge = Math.Tan(Angle / 2.0)
                        With Polyline
                            For Index As Integer = 0 To Point2dCollection.Count - 1
                                Point2d = Point2dCollection.Item(Index)
                                If Index = 0 Then
                                    .AddVertexAt(Index, Point2d, Bulge, 0, 0)
                                Else
                                    .AddVertexAt(Index, Point2d, 0, 0, 0)
                                End If
                            Next
                            .Closed = False
                        End With
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        BlockTableRecord.AppendEntity(Polyline)
                        Transaction.AddNewlyCreatedDBObject(Polyline, True)
                        Transaction.Commit()
                        Return Polyline
                    Catch
                        Transaction.Abort()
                        Return Nothing
                    End Try
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Junta entidades
        ''' </summary>
        ''' <param name="DbObjectCollection">Coleção de entidades a serem unidas</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Join(DbObjectCollection As DBObjectCollection) As Entity
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim DbBase As Entity
            Dim DbJoin As Entity
            Dim Curve As Curve
            Dim Point3dCollection As Point3dCollection
            Dim LastValue As Integer = 0
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    Try
                        While DbObjectCollection.Count <> 1
                            LastValue = DbObjectCollection.Count
                            DbBase = Transaction.GetObject(DbObjectCollection.Item(0).ObjectId, OpenMode.ForRead)
                            Select Case DbBase.GetType.Name
                                Case "Line"
                                    Curve = DbBase
                                    Point3dCollection = New Point3dCollection
                                    Point3dCollection.Add(Curve.StartPoint)
                                    Point3dCollection.Add(Curve.EndPoint)
                                    DbObjectCollection.Item(0) = Engine2.Drawing.DrawPolyline(Point3dCollection, False, Curve.Layer)
                                    Curve.UpgradeOpen()
                                    Curve.Erase()
                                    DbBase = Transaction.GetObject(DbObjectCollection.Item(0).ObjectId, OpenMode.ForRead)
                            End Select
                            DbJoin = Transaction.GetObject(DbObjectCollection.Item(1).ObjectId, OpenMode.ForRead)
                            DbBase.UpgradeOpen()
                            CType(DbBase, Polyline).JoinEntity(DbJoin)
                            DbJoin.UpgradeOpen()
                            DbJoin.Erase()
                            DbObjectCollection.Remove(DbJoin)
                            If DbObjectCollection.Count = LastValue And LastValue <> 0 Then
                                Exit While
                            End If
                        End While
                        Transaction.Commit()
                        Return DbObjectCollection.Item(0)
                    Catch
                        Transaction.Abort()
                        Return Nothing
                    End Try
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Junta entidades
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DbObjectCollection">Coleção de entidades a serem unidas</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Join(Transaction As Transaction, DbObjectCollection As DBObjectCollection) As Entity
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim DbBase As Entity
            Dim DbJoin As Entity
            Dim Curve As Curve
            Dim Point3dCollection As Point3dCollection
            Dim LastValue As Integer = 0
            Dim Polyline As Polyline
            Try
                While DbObjectCollection.Count <> 1
                    LastValue = DbObjectCollection.Count
                    DbBase = Transaction.GetObject(DbObjectCollection.Item(0).ObjectId, OpenMode.ForRead)
                    Select Case DbBase.GetType.Name
                        Case "Line"
                            Curve = DbBase
                            Point3dCollection = New Point3dCollection
                            Point3dCollection.Add(Curve.StartPoint)
                            Point3dCollection.Add(Curve.EndPoint)
                            DbObjectCollection.Item(0) = Engine2.Drawing.DrawPolyline(Point3dCollection, False, Curve.Layer)
                            Curve.UpgradeOpen()
                            Curve.Erase()
                            DbBase = Transaction.GetObject(DbObjectCollection.Item(0).ObjectId, OpenMode.ForRead)
                        Case "Arc"
                            Curve = DbBase
                            Polyline = Engine2.EntityInteration.GetPolylineFromArc(DbBase)
                            DbObjectCollection.Item(0) = Polyline
                            Curve.UpgradeOpen()
                            Curve.Erase()
                            DbBase = Transaction.GetObject(DbObjectCollection.Item(0).ObjectId, OpenMode.ForRead)
                    End Select
                    DbJoin = Transaction.GetObject(DbObjectCollection.Item(1).ObjectId, OpenMode.ForRead)
                    DbBase.UpgradeOpen()
                    CType(DbBase, Polyline).JoinEntity(DbJoin)
                    DbJoin.UpgradeOpen()
                    DbJoin.Erase()
                    DbObjectCollection.Remove(DbJoin)
                    If DbObjectCollection.Count = LastValue And LastValue <> 0 Then
                        Exit While
                    End If
                End While
                Return DbObjectCollection.Item(0)
            Catch
                Return Nothing
            End Try
        End Function

#Region "Função GetEntityIntersections e classes associadas"

        ''' <summary>
        ''' Obtem as intersecções e entidades associadas
        ''' </summary>
        ''' <param name="Entity">Entidade a ser analisada</param>
        ''' <param name="IntersectType">Tipo de captura de intersecção</param>
        ''' <param name="TypedValueCollection">Filtro de seleção</param>
        ''' <param name="IgnoreIds">ID´s que serão desconsideradas na seleção</param>
        ''' <param name="StepSurfaceCurve">Número de segmentos ao longo da entidade para detecção de intersecções em curvas</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetEntityIntersections(Entity As Entity, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.OnBothOperands, Optional TypedValueCollection As List(Of TypedValue) = Nothing, Optional IgnoreIds As ObjectIdCollection = Nothing, Optional StepSurfaceCurve As Integer = 250) As IntersectionEntityCollection
            GetEntityIntersections = New IntersectionEntityCollection(Entity)
            Dim ObjectIdCollection As New ObjectIdCollection
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Curve As Curve
            Dim Point3dcollection As Point3dCollection
            Dim Point3dcollection2 As Point3dCollection
            Dim Point3d As Point3d
            Dim Lenght As Double
            Dim [Step] As Double
            Dim EntityTypeString As String = ""
            Dim Item As Integer = 0
            Dim PromptSelectionResult As PromptSelectionResult
            Dim IntersectEntity As Entity = Nothing
            Dim DBObjectCollection As DBObjectCollection = GetSubEntitys(Entity, True)
            Dim Circle As Circle
            Dim Ellipse As Ellipse
            Dim BlockReference As BlockReference
            Dim DBPoint As DBPoint
            Dim DBObjectCollection2 As DBObjectCollection
            Dim Create As Boolean
            Dim Curve2 As Curve
            If DBObjectCollection.Count > 0 Then
                For Each DBObject As DBObject In DBObjectCollection
                    If Engine2.EntityInteration.IsCurved(DBObject) = True Then
                        Point3dcollection = New Point3dCollection()
                        Select Case DBObject.GetType.Name
                            Case "Line"
                                Curve = DBObject
                                With Point3dcollection
                                    .Add(Curve.StartPoint)
                                    .Add(Curve.EndPoint)
                                End With
                            Case "Circle"
                                Circle = DBObject
                                Curve = Engine2.EntityInteration.OffSet(Circle, 0.01)
                                Lenght = Curve.GetDistanceAtParameter(Curve.EndParam) - Curve.GetDistanceAtParameter(Curve.StartParam)
                                [Step] = (Lenght / StepSurfaceCurve)
                                For CurvePosition As Double = 0 To Lenght Step [Step]
                                    Point3d = Curve.GetPointAtDist(CurvePosition)
                                    Point3dcollection.Add(Curve.GetClosestPointTo(Point3d, False))
                                Next
                                If Point3dcollection.Item(0).Equals(Point3dcollection.Item(Point3dcollection.Count - 1)) = False Then
                                    Point3dcollection.Add(Point3dcollection.Item(0))
                                End If
                            Case "Ellipse"
                                Ellipse = DBObject
                                Curve = Engine2.EntityInteration.OffSet(Ellipse, 0.01)
                                Lenght = Curve.GetDistanceAtParameter(Curve.EndParam) - Curve.GetDistanceAtParameter(Curve.StartParam)
                                [Step] = (Lenght / StepSurfaceCurve)
                                For CurvePosition As Double = 0 To Lenght Step [Step]
                                    Point3d = Curve.GetPointAtDist(CurvePosition)
                                    Point3dcollection.Add(Curve.GetClosestPointTo(Point3d, False))
                                Next
                                If Ellipse.Closed = True Then
                                    If Point3dcollection.Item(0).Equals(Point3dcollection.Item(Point3dcollection.Count - 1)) = False Then
                                        Point3dcollection.Add(Point3dcollection.Item(0))
                                    End If
                                End If
                            Case Else
                                Curve = DBObject
                                Lenght = Curve.GetDistanceAtParameter(Curve.EndParam) - Curve.GetDistanceAtParameter(Curve.StartParam)
                                [Step] = (Lenght / StepSurfaceCurve)
                                For CurvePosition As Double = 0 To Lenght Step [Step]
                                    Point3d = Curve.GetPointAtDist(CurvePosition)
                                    Point3dcollection.Add(Curve.GetClosestPointTo(Point3d, False))
                                Next
                        End Select
                        If IsNothing(TypedValueCollection) = False Then
                            PromptSelectionResult = Editor.SelectFence(Point3dcollection, New SelectionFilter(TypedValueCollection.ToArray))
                        Else
                            PromptSelectionResult = Editor.SelectFence(Point3dcollection)
                        End If
                        If PromptSelectionResult.Status = PromptStatus.OK Then
                            Point3dcollection.Clear()
                            For Each ObjectId As ObjectId In PromptSelectionResult.Value.GetObjectIds
                                If ObjectId <> Entity.ObjectId Then
                                    IntersectEntity = Engine2.ConvertObject.ObjectIDToEntity(ObjectId)
                                    IntersectEntity.IntersectWith(Entity, IntersectType, Point3dcollection, IntPtr.Zero, IntPtr.Zero)
                                    Select Case IntersectEntity.GetType.Name
                                        Case "Line", "Arc", "Polyline", "Spline"
                                            Curve2 = IntersectEntity
                                            Curve = Entity
                                            If Point3dcollection.Contains(Curve.StartPoint) = True Then
                                                Point3dcollection.Remove(Curve.StartPoint)
                                            End If
                                            If Point3dcollection.Contains(Curve.EndPoint) = True Then
                                                Point3dcollection.Remove(Curve.EndPoint)
                                            End If
                                            If Point3dcollection.Count > 0 Then
                                                If IsNothing(IgnoreIds) = True Then
                                                    Point3dcollection2 = New Point3dCollection
                                                    For Each Point3d In Point3dcollection
                                                        Point3dcollection2.Add(Point3d)
                                                        If EntityInteration.PointAtEntity(Curve, Curve2.StartPoint) = True Then
                                                            If Point3dcollection.Contains(Curve2.StartPoint) = False Then
                                                                Point3dcollection2.Add(Curve2.StartPoint)
                                                            End If
                                                        End If
                                                        If EntityInteration.PointAtEntity(Curve, Curve2.EndPoint) = True Then
                                                            If Point3dcollection.Contains(Curve2.EndPoint) = False Then
                                                                Point3dcollection2.Add(Curve2.EndPoint)
                                                            End If
                                                        End If
                                                    Next
                                                    GetEntityIntersections.Add(IntersectEntity, Point3dcollection2)
                                                Else
                                                    If IgnoreIds.Contains(IntersectEntity.ObjectId) = False Then
                                                        Point3dcollection2 = New Point3dCollection
                                                        For Each Point3d In Point3dcollection
                                                            Point3dcollection2.Add(Point3d)
                                                            If EntityInteration.PointAtEntity(Curve, Curve2.StartPoint) = True Then
                                                                If Point3dcollection.Contains(Curve2.StartPoint) = False Then
                                                                    Point3dcollection2.Add(Curve2.StartPoint)
                                                                End If
                                                            End If
                                                            If EntityInteration.PointAtEntity(Curve, Curve2.EndPoint) = True Then
                                                                If Point3dcollection.Contains(Curve2.EndPoint) = False Then
                                                                    Point3dcollection2.Add(Curve2.EndPoint)
                                                                End If
                                                            End If
                                                        Next
                                                        GetEntityIntersections.Add(IntersectEntity, Point3dcollection2)
                                                    End If
                                                End If
                                            End If
                                        Case "Circle"
                                            Circle = Entity
                                            If Point3dcollection.Contains(Circle.Center) = True Then
                                                Point3dcollection.Remove(Circle.Center)
                                            End If
                                            If Point3dcollection.Count > 0 Then
                                                If IsNothing(IgnoreIds) = True Then
                                                    Point3dcollection2 = New Point3dCollection
                                                    For Each Point3d In Point3dcollection
                                                        Point3dcollection2.Add(Point3d)
                                                    Next
                                                    GetEntityIntersections.Add(IntersectEntity, Point3dcollection2)
                                                Else
                                                    If IgnoreIds.Contains(IntersectEntity.ObjectId) = False Then
                                                        Point3dcollection2 = New Point3dCollection
                                                        For Each Point3d In Point3dcollection
                                                            Point3dcollection2.Add(Point3d)
                                                        Next
                                                        GetEntityIntersections.Add(IntersectEntity, Point3dcollection2)
                                                    End If
                                                End If
                                            End If
                                        Case "Ellipse"
                                            Ellipse = Entity
                                            If Point3dcollection.Contains(Ellipse.Center) = True Then
                                                Point3dcollection.Remove(Ellipse.Center)
                                            End If
                                            If Point3dcollection.Count > 0 Then
                                                If IsNothing(IgnoreIds) = True Then
                                                    Point3dcollection2 = New Point3dCollection
                                                    For Each Point3d In Point3dcollection
                                                        Point3dcollection2.Add(Point3d)
                                                    Next
                                                    GetEntityIntersections.Add(IntersectEntity, Point3dcollection2)
                                                Else
                                                    If IgnoreIds.Contains(IntersectEntity.ObjectId) = False Then
                                                        Point3dcollection2 = New Point3dCollection
                                                        For Each Point3d In Point3dcollection
                                                            Point3dcollection2.Add(Point3d)
                                                        Next
                                                        GetEntityIntersections.Add(IntersectEntity, Point3dcollection2)
                                                    End If
                                                End If
                                            End If
                                        Case "BlockReference"
                                            BlockReference = Entity
                                            If Point3dcollection.Contains(BlockReference.Position) = True Then
                                                Point3dcollection.Remove(BlockReference.Position)
                                            End If
                                            If Point3dcollection.Count > 0 Then
                                                If IsNothing(IgnoreIds) = True Then
                                                    Point3dcollection2 = New Point3dCollection
                                                    For Each Point3d In Point3dcollection
                                                        Point3dcollection2.Add(Point3d)
                                                    Next
                                                    GetEntityIntersections.Add(IntersectEntity, Point3dcollection2)
                                                Else
                                                    If IgnoreIds.Contains(IntersectEntity.ObjectId) = False Then
                                                        Point3dcollection2 = New Point3dCollection
                                                        For Each Point3d In Point3dcollection
                                                            Point3dcollection2.Add(Point3d)
                                                        Next
                                                        GetEntityIntersections.Add(IntersectEntity, Point3dcollection2)
                                                    End If
                                                End If
                                            End If
                                        Case "DBPoint"
                                            DBPoint = Entity
                                            If Point3dcollection.Contains(DBPoint.Position) = True Then
                                                Point3dcollection.Remove(DBPoint.Position)
                                            End If
                                            If Point3dcollection.Count > 0 Then
                                                If IsNothing(IgnoreIds) = True Then
                                                    Point3dcollection2 = New Point3dCollection
                                                    For Each Point3d In Point3dcollection
                                                        Point3dcollection2.Add(Point3d)
                                                    Next
                                                    GetEntityIntersections.Add(IntersectEntity, Point3dcollection2)
                                                Else
                                                    If IgnoreIds.Contains(IntersectEntity.ObjectId) = False Then
                                                        Point3dcollection2 = New Point3dCollection
                                                        For Each Point3d In Point3dcollection
                                                            Point3dcollection2.Add(Point3d)
                                                        Next
                                                        GetEntityIntersections.Add(IntersectEntity, Point3dcollection2)
                                                    End If
                                                End If
                                            End If
                                        Case Else
                                            Throw New System.Exception("O tipo '" & IntersectEntity.GetType.Name & "' não é válido para GetEntityIntersections.")
                                    End Select
                                End If
                            Next
                        End If
                        If Engine2.EntityInteration.IsCurved(DBObject) = True Then
                            With TypedValueCollection
                                .Clear()
                                .Add(New TypedValue(DxfCode.Operator, "<OR"))
                                .Add(New TypedValue(DxfCode.Start, "LINE"))
                                .Add(New TypedValue(DxfCode.Start, "LWPOLYLINE"))
                                .Add(New TypedValue(DxfCode.Start, "SPLINE"))
                                .Add(New TypedValue(DxfCode.Start, "ARC"))
                                .Add(New TypedValue(DxfCode.Operator, "OR>"))
                            End With
                            Curve = DBObject
                            Create = True
                            DBObjectCollection2 = Engine2.AcadInterface.SelectAtPoint(Curve.StartPoint, TypedValueCollection, New ObjectIdCollection({DBObject.ObjectId}), AcadInterface.eSelectAtPointRules.ThatBeginOrEndPoint)
                            For Each DBObject2 As DBObject In DBObjectCollection2
                                If IsNothing(IgnoreIds) = True Then
                                    For Index As Integer = 0 To GetEntityIntersections.Count - 1
                                        If GetEntityIntersections.Item(Index).Entity.Handle = DBObject.Handle Then
                                            If GetEntityIntersections.IntersectionPoints3d.Contains(Curve.StartPoint) = False Then
                                                GetEntityIntersections.IntersectionPoints3d.Add(Curve.StartPoint)
                                                Create = False
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    If Create = True Then
                                        GetEntityIntersections.Add(DBObject2, New List(Of Point3d)({Curve.StartPoint}))
                                    End If
                                Else
                                    If IgnoreIds.Contains(DBObject2.ObjectId) = False Then
                                        For Index As Integer = 0 To GetEntityIntersections.Count - 1
                                            If GetEntityIntersections.Item(Index).Entity.Handle = DBObject.Handle Then
                                                If GetEntityIntersections.IntersectionPoints3d.Contains(Curve.StartPoint) = False Then
                                                    GetEntityIntersections.IntersectionPoints3d.Add(Curve.StartPoint)
                                                    Create = False
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                        If Create = True Then
                                            GetEntityIntersections.Add(DBObject2, New List(Of Point3d)({Curve.StartPoint}))
                                        End If
                                    End If
                                End If
                            Next
                            Create = True
                            DBObjectCollection2 = Engine2.AcadInterface.SelectAtPoint(Curve.EndPoint, TypedValueCollection, New ObjectIdCollection({DBObject.ObjectId}), AcadInterface.eSelectAtPointRules.ThatBeginOrEndPoint)
                            For Each DBObject2 As DBObject In DBObjectCollection2
                                If IsNothing(IgnoreIds) = True Then
                                    For Index As Integer = 0 To GetEntityIntersections.Count - 1
                                        If GetEntityIntersections.Item(Index).Entity.Handle = DBObject.Handle Then
                                            If GetEntityIntersections.IntersectionPoints3d.Contains(Curve.EndPoint) = False Then
                                                GetEntityIntersections.IntersectionPoints3d.Add(Curve.EndPoint)
                                                Create = False
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    If Create = True Then
                                        GetEntityIntersections.Add(DBObject2, New List(Of Point3d)({Curve.EndPoint}))
                                    End If
                                Else
                                    If IgnoreIds.Contains(DBObject2.ObjectId) = False Then
                                        For Index As Integer = 0 To GetEntityIntersections.Count - 1
                                            If GetEntityIntersections.Item(Index).Entity.Handle = DBObject.Handle Then
                                                If GetEntityIntersections.IntersectionPoints3d.Contains(Curve.EndPoint) = False Then
                                                    GetEntityIntersections.IntersectionPoints3d.Add(Curve.EndPoint)
                                                    Create = False
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                        If Create = True Then
                                            GetEntityIntersections.Add(DBObject2, New List(Of Point3d)({Curve.EndPoint}))
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
            Return GetEntityIntersections
        End Function




        ''' <summary>
        ''' Determina se uma entidade pode ser curva
        ''' </summary>
        ''' <param name="Entity">Entity</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsCurved(Entity As Entity) As Boolean
            Try
                Dim Curve As Curve = Entity
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Coleção de entidades e pontos de intersecção (Retorno de GetIntersection)
        ''' </summary>
        ''' <remarks></remarks>
        Public NotInheritable Class IntersectionEntityCollection : Inherits List(Of IntersectionEntity)

            ''' <summary>
            ''' Construtor
            ''' </summary>
            ''' <param name="EntityBase">Entidade base</param>
            ''' <remarks></remarks>
            Public Sub New(EntityBase As Entity)
                Me._EntityBase = EntityBase
            End Sub

            ''' <summary>
            ''' Entidade base geradora do cálculo
            ''' </summary>
            ''' <remarks></remarks>
            Private _EntityBase As Entity
            Public ReadOnly Property EntityBase As Entity
                Get
                    Return Me._EntityBase
                End Get
            End Property

            ''' <summary>
            ''' Pontos de intersecção 3d
            ''' </summary>
            ''' <remarks></remarks>
            Public ReadOnly Property IntersectionPoints3d As List(Of Point3d)
                Get
                    IntersectionPoints3d = New List(Of Point3d)
                    For Index As Integer = 0 To MyBase.Count - 1
                        For Each Point As Object In MyBase.Item(Index).IntersectionPoints
                            If IntersectionPoints3d.Contains(Point) = False Then
                                IntersectionPoints3d.Add(Point)
                            End If
                        Next
                    Next
                    If IntersectionPoints3d.Count > 0 Then
                        Return IntersectionPoints3d
                    Else
                        Return Nothing
                    End If
                End Get
            End Property

            ''' <summary>
            ''' Pontos de intersecção 2d
            ''' </summary>
            ''' <remarks></remarks>
            Public ReadOnly Property IntersectionPoints2d As List(Of Point2d)
                Get
                    IntersectionPoints2d = New List(Of Point2d)
                    Dim Point3d As Point3d
                    For Index As Integer = 0 To MyBase.Count - 1
                        For Each Point As Object In MyBase.Item(Index).IntersectionPoints
                            Point3d = Engine2.Geometry.Point2dToPoint3d(Point)
                            If IntersectionPoints3d.Contains(Point3d) = False Then
                                IntersectionPoints3d.Add(Point3d)
                            End If
                        Next
                    Next
                    If IntersectionPoints2d.Count > 0 Then
                        Return IntersectionPoints2d
                    Else
                        Return Nothing
                    End If
                End Get
            End Property

            ''' <summary>
            ''' Coleção de entidades
            ''' </summary>
            ''' <remarks></remarks>
            Public ReadOnly Property EntityCollection As List(Of Entity)
                Get
                    EntityCollection = New List(Of Entity)
                    For Index As Integer = 0 To MyBase.Count - 1
                        If EntityCollection.Contains(MyBase.Item(Index).Entity) = False Then
                            EntityCollection.Add(MyBase.Item(Index).Entity)
                        End If
                    Next
                    If EntityCollection.Count > 0 Then
                        Return EntityCollection
                    Else
                        Return Nothing
                    End If
                End Get
            End Property

            ''' <summary>
            ''' Coleção de ID´s das entidades
            ''' </summary>
            ''' <remarks></remarks>
            Public ReadOnly Property ObjectIDCollection As Object
                Get
                    ObjectIDCollection = New ObjectIdCollection
                    For Index As Integer = 0 To MyBase.Count - 1
                        If ObjectIDCollection.Contains(MyBase.Item(Index).Entity.ObjectId) = False Then
                            ObjectIDCollection.Add(MyBase.Item(Index).Entity.ObjectId)
                        End If
                    Next
                    If ObjectIDCollection.Count > 0 Then
                        Return ObjectIDCollection
                    Else
                        Return Nothing
                    End If
                End Get
            End Property

            ''' <summary>
            ''' Adicionar
            ''' </summary>
            ''' <remarks></remarks>
            Public Overloads Sub Add(Entity As Entity, IntersectionPoints As Object)
                MyBase.Add(New IntersectionEntity(Entity, IntersectionPoints))
            End Sub

            ''' <summary>
            ''' Item de IntersectionEntityCollection
            ''' </summary>
            ''' <remarks></remarks>
            Public NotInheritable Class IntersectionEntity

                ''' <summary>
                ''' Construtor
                ''' </summary>
                ''' <param name="Entity">Entidade</param>
                ''' <param name="IntersectionPoints">Coleção de pontos de intersecção</param>
                ''' <remarks></remarks>
                Public Sub New(Entity As Entity, IntersectionPoints As Object)
                    Me._IntersectionPoints = New List(Of Point3d)
                    Me._Entity = Entity
                    For Each Point As Object In IntersectionPoints
                        Select Case Point.GetType.Name
                            Case "Point2d"
                                Me._IntersectionPoints.Add(Engine2.Geometry.Point2dToPoint3d(Point))
                            Case "Point3d"
                                Me._IntersectionPoints.Add(Point)
                        End Select
                    Next
                End Sub

                ''' <summary>
                ''' Ponto de intersecção
                ''' </summary>
                ''' <remarks></remarks>
                Private _IntersectionPoints As List(Of Point3d)
                Public ReadOnly Property IntersectionPoints As List(Of Point3d)
                    Get
                        Return Me._IntersectionPoints
                    End Get
                End Property

                ''' <summary>
                ''' Entidade
                ''' </summary>
                ''' <remarks></remarks>
                Private _Entity As Entity
                Public ReadOnly Property Entity As Entity
                    Get
                        Return Me._Entity
                    End Get
                End Property

            End Class

        End Class

#End Region

    End Class

End Namespace



