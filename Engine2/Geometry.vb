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
Imports Autodesk.AutoCAD.Geometry
Imports System.Reflection
Imports ElectricTools2020.Engine2.Geometry
Imports ElectricTools2020.Engine2.EntityInteration

Namespace Engine2

    ''' <summary>
    ''' Classe de geometria do sistema
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Geometry

        ''' <summary>
        ''' Formato angular
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eAngleFormat

            ''' <summary>
            ''' Radiano
            ''' </summary>
            ''' <remarks></remarks>
            Radian = 0

            ''' <summary>
            ''' Grau
            ''' </summary>
            ''' <remarks></remarks>
            Degrees = 1

        End Enum

        ''' <summary>
        ''' Eixo
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eAxis

            ''' <summary>
            ''' Eixo X
            ''' </summary>
            ''' <remarks></remarks>
            X = 0

            ''' <summary>
            ''' Eixo Y
            ''' </summary>
            ''' <remarks></remarks>
            Y = 1

            ''' <summary>
            ''' Eixo Z
            ''' </summary>
            ''' <remarks></remarks>
            Z = 2

        End Enum

        ''' <summary>
        ''' Determina se um número é par
        ''' </summary>
        ''' <param name="Number"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsEven(Number As Double) As Boolean
            If Number Mod 2 <> 0 Then
                Return False
            Else
                Return True
            End If
        End Function

        ''' <summary>
        ''' Retorna o ângulo reverso
        ''' </summary>
        ''' <param name="DegreeAngle">Ângulo em graus</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ReverseDegreeAngle(DegreeAngle As Double) As Double
            Return Engine2.Geometry.AngleAdd(DegreeAngle, 180)
        End Function

        ''' <summary>
        ''' Remove zeros a direita
        ''' </summary>
        ''' <param name="Value">Valor</param>
        ''' <returns>Double</returns>
        ''' <remarks></remarks>
        Public Shared Function SupressZeroRight(Value As Double) As Double
            Return Double.Parse(Value).ToString()
        End Function

        ''' <summary>
        ''' Modifica valores em um dos eixos da coordenada
        ''' </summary>
        ''' <param name="Point">Ponto a ser modificado</param>
        ''' <param name="Axis">Eixo a ser modificado</param>
        ''' <param name="Value">Valor a ser adicionado ou removido</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ChangePoint(Point As Point3d, Axis As eAxis, Value As Double) As Point3d
            Select Case Axis
                Case eAxis.X
                    Return New Point3d((Point.X + Value), Point.Y, Point.Z)
                Case eAxis.Y
                    Return New Point3d(Point.X, (Point.Y + Value), Point.Z)
                Case eAxis.Z
                    Return New Point3d(Point.X, Point.Y, (Point.Z + Value))
            End Select
        End Function

        ''' <summary>
        ''' Reverte o valor numérico positivo para negativo e negativo para positivo
        ''' </summary>
        ''' <param name="Value">Valor</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ReverseSymbol(Value As Double) As Double
            If Value > 0 Then
                Return (Value * -1)
            ElseIf Value < 0 Then
                Return (Value / -1)
            Else
                Return Value
            End If
        End Function

        ''' <summary>
        ''' Reverte o valor numérico positivo para negativo e negativo para positivo
        ''' </summary>
        ''' <param name="Value">Valor</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ReverseSymbol(Value As Integer) As Integer
            If Value > 0 Then
                Return (Value * -1)
            ElseIf Value < 0 Then
                Return (Value / -1)
            Else
                Return Value
            End If
        End Function

        ''' <summary>
        ''' Obtem o percentual referente a coordenada em relação ao comprimento do segmento com base em seu ponto inicial
        ''' </summary>
        ''' <param name="Curve"></param>
        ''' <param name="Point3d"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetPercentageAtPoint(Curve As Curve, Point3d As Point3d) As Double
            Return ((Curve.GetDistAtPoint(Point3d) / Engine2.Geometry.GetLength(Curve)) * 100)
        End Function

        ''' <summary>
        ''' Obtem a coordenada em relação ao percentual da curva com base em seu ponto inicial
        ''' </summary>
        ''' <param name="Curve">Curve</param>
        ''' <param name="Percentage">Percentual</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetPointAtPercentage(Curve As Curve, Percentage As Double) As Point3d
            Return Curve.GetPointAtDist((Engine2.Geometry.GetLength(Curve) * Percentage) / 100)
        End Function

        ''' <summary>
        ''' Calcula a hipotenusa do triangulo retângulo
        ''' </summary>
        ''' <param name="Side1">Tamanho do lado 1</param>
        ''' <param name="Side2">Tamanho do lado 2</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Hypotenuse(Side1 As Double, Side2 As Double) As Double
            Return Math.Sqrt((Side1 ^ 2) + (Side2 ^ 2))
        End Function

        ''' <summary>
        ''' Cálcula o cateto do triângulo retângulo
        ''' </summary>
        ''' <param name="Side1">Tamanho do lado 1</param>
        ''' <param name="Side2">Tamanho do lado 2</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Cateto(Side1 As Double, Side2 As Double) As Double
            Return Math.Sqrt(Math.Max((Side1 ^ 2), (Side2 ^ 2)) - Math.Min((Side1 ^ 2), (Side2 ^ 2)))
        End Function

        ''' <summary>
        ''' Retorna o comprimento da tangente de um arco
        ''' </summary>
        ''' <param name="DegreeAngle"></param>
        ''' <param name="Radius"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ArcTangent(DegreeAngle As Double, Radius As Double) As Double
            Return (Radius * Math.Tan(Engine2.Geometry.DegreeToRadian((CDbl(DegreeAngle) / 2))))
        End Function

        ''' <summary>
        ''' Obtem o centro de polylines 
        ''' </summary>
        ''' <param name="Point3dCollection">Coleção de pontos 3d (Vértices)</param>
        ''' <returns>Point2d\Nothing</returns>
        ''' <remarks></remarks>
        Public Shared Function GetCentroid(Point3dCollection As Point3dCollection) As Object
            Try
                Dim ReferencePoint As Point3d
                Dim RunningX As Double
                Dim RunningY As Double
                Dim SecondFactor As Double
                Dim Point1 As Point3d
                Dim Point2 As Point3d
                Dim X As Double
                Dim Y As Double
                Dim Z As Double = 0
                Dim Polyline As Polyline
                Dim Area As Double
                Dim Point2d As Point2d
                Dim Point2dCollection As New Point2dCollection
                For Each Point3d As Point3d In Point3dCollection
                    Point2d = Point3d.Convert2d(New Plane(Point3d.Origin, Vector3d.ZAxis))
                    Point2dCollection.Add(Point2d)
                Next
                Polyline = New Polyline(Point2dCollection.Count)
                With Polyline
                    For Index As Integer = 0 To Point2dCollection.Count - 1
                        Point2d = Point2dCollection.Item(Index)
                        .AddVertexAt(Index, Point2d, 0, 0, 0)
                    Next
                    .Closed = True
                End With
                Area = Polyline.Area
                Polyline.Dispose()
                For Vertex As Integer = 0 To Point3dCollection.Count - 2
                    Point1 = Point3dCollection.Item(Vertex)
                    Point2 = Point3dCollection.Item(Vertex + 1)
                    If Vertex = 0 Then
                        ReferencePoint = Point1
                    End If
                    SecondFactor = (Point1.X - ReferencePoint.X) * (Point2.Y - ReferencePoint.Y) - (Point2.X - ReferencePoint.X) * (Point1.Y - ReferencePoint.Y)
                    RunningX += ((Point1.X - ReferencePoint.X) + (Point2.X - ReferencePoint.X)) * SecondFactor
                    RunningY += ((Point1.Y - ReferencePoint.Y) + (Point2.Y - ReferencePoint.Y)) * SecondFactor
                Next
                X = RunningX / (6 * Area) + ReferencePoint.X
                Y = RunningY / (6 * Area) + ReferencePoint.Y
                Return New Point3d(X, Y, Z)
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem o centro de polylines 
        ''' </summary>
        ''' <param name="Point2dCollection">Coleção de pontos 2d (Vértices)</param>
        ''' <returns>Point2d\Nothing</returns>
        ''' <remarks></remarks>
        Public Shared Function GetCentroid(Point2dCollection As Point2dCollection) As Object
            Try
                Dim ReferencePoint As Point2d
                Dim RunningX As Double
                Dim RunningY As Double
                Dim SecondFactor As Double
                Dim Point1 As Point2d
                Dim Point2 As Point2d
                Dim X As Double
                Dim Y As Double
                Dim Polyline As Polyline
                Dim Area As Double
                Dim Point2d As Point2d
                Polyline = New Polyline(Point2dCollection.Count)
                With Polyline
                    For Index As Integer = 0 To Point2dCollection.Count - 1
                        Point2d = Point2dCollection.Item(Index)
                        .AddVertexAt(Index, Point2d, 0, 0, 0)
                    Next
                    .Closed = True
                End With
                Area = Polyline.Area
                Polyline.Dispose()
                For Vertex As Integer = 0 To Point2dCollection.Count - 2
                    Point1 = Point2dCollection.Item(Vertex)
                    Point2 = Point2dCollection.Item(Vertex + 1)
                    If Vertex = 0 Then
                        ReferencePoint = Point1
                    End If
                    SecondFactor = (Point1.X - ReferencePoint.X) * (Point2.Y - ReferencePoint.Y) - (Point2.X - ReferencePoint.X) * (Point1.Y - ReferencePoint.Y)
                    RunningX += ((Point1.X - ReferencePoint.X) + (Point2.X - ReferencePoint.X)) * SecondFactor
                    RunningY += ((Point1.Y - ReferencePoint.Y) + (Point2.Y - ReferencePoint.Y)) * SecondFactor
                Next
                X = RunningX / (6 * Area) + ReferencePoint.X
                Y = RunningY / (6 * Area) + ReferencePoint.Y
                Return New Point2d(X, Y)
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem a intersecção para um determinado ângulo com base em um ponto
        ''' </summary>
        ''' <returns>Point3d\Nothing</returns>
        ''' <remarks></remarks>
        Public Shared Function GetIntersectionAtAngle(LineSegment3d As LineSegment3d, PointReference As Point3d, DegreesAngle As Double) As Object
            Dim PerpendicularInfor As PerpendicularInfor
            Dim Angle As Double
            Dim Point3dCollection As Point3dCollection
            PerpendicularInfor = Engine2.Geometry.GetPerpendicular(LineSegment3d.StartPoint, LineSegment3d.EndPoint, PointReference)
            If IsNothing(PerpendicularInfor) = False Then
                Angle = Engine2.Geometry.GetAngle(PointReference, PerpendicularInfor.PerpendicularPoint, eAngleFormat.Degrees)
                Point3dCollection = Engine2.Geometry.VirtualIntersection(PointReference, Engine2.Geometry.AngleAdd(Angle, -DegreesAngle), LineSegment3d, Intersect.ExtendBoth)
                If Point3dCollection.Count > 0 Then
                    Return Point3dCollection.Item(0)
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Obtem a intersecção para um determinado ângulo e suas extremidades com base em um ponto
        ''' </summary>
        ''' <returns>Point3d\Nothing</returns>
        ''' <remarks></remarks>
        Public Shared Function GetIntersectionAtDerivedAngle(LineSegment3d As LineSegment3d, PointReference As Point3d, DegreesAngle As Double) As Object
            Dim PerpendicularInfor As PerpendicularInfor
            Dim Angle As Double
            Dim Point3dCollection As Point3dCollection
            PerpendicularInfor = Engine2.Geometry.GetPerpendicular(LineSegment3d.StartPoint, LineSegment3d.EndPoint, PointReference)
            If IsNothing(PerpendicularInfor) = False Then
                If PerpendicularInfor.PerpendicularPoint.DistanceTo(LineSegment3d.MidPoint) <= Tolerance.Global.EqualPoint Then
                    Return Nothing
                Else
                    Angle = Engine2.Geometry.GetAngle(PointReference, PerpendicularInfor.PerpendicularPoint, eAngleFormat.Degrees)
                    If PerpendicularInfor.VirtualPerpendicular = True Then
                        If Engine2.Geometry.GetAngle(PerpendicularInfor.ProjectionBasePoint, PerpendicularInfor.PerpendicularPoint, PointReference, eAngleFormat.Degrees) <= 180 Then
                            Point3dCollection = Engine2.Geometry.VirtualIntersection(PointReference, Engine2.Geometry.AngleAdd(Angle, (90 - DegreesAngle)), LineSegment3d, Intersect.ExtendBoth)
                        Else
                            Point3dCollection = Engine2.Geometry.VirtualIntersection(PointReference, Engine2.Geometry.AngleAdd(Angle, -(90 - DegreesAngle)), LineSegment3d, Intersect.ExtendBoth)
                        End If
                    Else
                        If Engine2.Geometry.GetAngle(PerpendicularInfor.ProjectionBasePoint, PerpendicularInfor.PerpendicularPoint, PointReference, eAngleFormat.Degrees) <= 180 Then
                            Point3dCollection = Engine2.Geometry.VirtualIntersection(PointReference, Engine2.Geometry.AngleAdd(Angle, -(90 - DegreesAngle)), LineSegment3d, Intersect.ExtendBoth)
                        Else
                            Point3dCollection = Engine2.Geometry.VirtualIntersection(PointReference, Engine2.Geometry.AngleAdd(Angle, (90 - DegreesAngle)), LineSegment3d, Intersect.ExtendBoth)
                        End If
                    End If
                    If Point3dCollection.Count > 0 Then
                        Return Point3dCollection.Item(0)
                    Else
                        Return Nothing
                    End If
                End If
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Retorna o ponto 2d vértice de um arco
        ''' </summary>
        ''' <returns>Point3d\Nothing</returns>
        ''' <remarks></remarks>
        Public Shared Function VertexOfArcTo2dPoint(Arc As Arc) As Object
            Return Engine2.Geometry.Point3dToPoint2d(VertexOfArcTo3dPoint(Arc))
        End Function

        ''' <summary>
        ''' Retorna o ponto 3d vértice de um arco
        ''' </summary>
        ''' <returns>Point3d\Nothing</returns>
        ''' <remarks></remarks>
        Public Shared Function VertexOfArcTo3dPoint(Arc As Arc) As Object
            If (360 - Engine2.Geometry.GetAngle(Arc.Center, Arc.StartPoint, Arc.EndPoint, eAngleFormat.Degrees)) < 180 Then
                Dim DegreeAngle01 As Double
                Dim DegreeAngle02 As Double
                Dim Line01 As LineSegment3d
                Dim Line02 As LineSegment3d
                DegreeAngle01 = Engine2.Geometry.GetAngle(Arc.Center, Arc.StartPoint, eAngleFormat.Degrees) + 90
                DegreeAngle02 = Engine2.Geometry.GetAngle(Arc.Center, Arc.EndPoint, eAngleFormat.Degrees) - 90
                Line01 = New LineSegment3d(Arc.StartPoint, Engine2.Geometry.PolarPoint3d(Arc.StartPoint, DegreeAngle01, (Arc.Length / 4)))
                Line02 = New LineSegment3d(Arc.EndPoint, Engine2.Geometry.PolarPoint3d(Arc.EndPoint, DegreeAngle02, (Arc.Length / 4)))
                Return Engine2.Geometry.VirtualIntersection(Line01, Line02, , Intersect.ExtendBoth).Item(0)
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Retorna o ponto com a maior distância em relação ao ponto de referência
        ''' </summary>
        ''' <param name="Point3dCollection">Coleção de pontos 3d</param>
        ''' <returns>Point3d</returns>
        ''' <remarks></remarks>
        Public Shared Function GetFarthestPoint(Point3dCollection As Point3dCollection, PointReference As Point3d) As Point3d
            GetFarthestPoint = Point3dCollection.Item(0)
            For Each Point3d As Point3d In Point3dCollection
                If Point3d.Equals(GetFarthestPoint) = False Then
                    If Point3d.DistanceTo(PointReference) > GetFarthestPoint.DistanceTo(PointReference) Then
                        GetFarthestPoint = Point3d
                    End If
                End If
            Next
            Return GetFarthestPoint
        End Function

        ''' <summary>
        ''' Retorna o ponto com a menor distância em relação ao ponto de referência
        ''' </summary>
        ''' <param name="Point3dCollection">Coleção de pontos 3d</param>
        ''' <returns>Point3d</returns>
        ''' <remarks></remarks>
        Public Shared Function GetClosestPoint(Point3dCollection As Point3dCollection, PointReference As Point3d) As Point3d
            GetClosestPoint = Point3dCollection.Item(0)
            For Each Point3d As Point3d In Point3dCollection
                If Point3d.Equals(GetClosestPoint) = False Then
                    If Point3d.DistanceTo(PointReference) < GetClosestPoint.DistanceTo(PointReference) Then
                        GetClosestPoint = Point3d
                    End If
                End If
            Next
            Return GetClosestPoint
        End Function

        ''' <summary>
        ''' Retorna a menor distância entre pontos de uma coleção
        ''' </summary>
        ''' <param name="Point3dCollection">Coleção de pontos 3d</param>
        ''' <returns>Point3d</returns>
        ''' <remarks></remarks>
        Public Shared Function MinDistanceFromPoints(Point3dCollection As Point3dCollection) As Double
            Dim Distance As Object = Nothing
            For Each PointBase As Point3d In Point3dCollection
                For Each PointCompare As Point3d In Point3dCollection
                    If PointBase.Equals(PointCompare) = False Then
                        If IsNothing(Distance) = True Then
                            Distance = PointCompare.DistanceTo(PointBase)
                        Else
                            If PointCompare.DistanceTo(PointBase) < Distance Then
                                Distance = PointCompare.DistanceTo(PointBase)
                            End If
                        End If
                    End If
                Next
            Next
            Return Distance
        End Function

        ''' <summary>
        ''' Retorna a maior distância entre pontos de uma coleção
        ''' </summary>
        ''' <param name="Point3dCollection">Coleção de pontos 3d</param>
        ''' <returns>Point3d</returns>
        ''' <remarks></remarks>
        Public Shared Function MaxDistanceFromPoints(Point3dCollection As Point3dCollection) As Double
            Dim Distance As Object = Nothing
            For Each PointBase As Point3d In Point3dCollection
                For Each PointCompare As Point3d In Point3dCollection
                    If PointBase.Equals(PointCompare) = False Then
                        If IsNothing(Distance) = True Then
                            Distance = PointCompare.DistanceTo(PointBase)
                        Else
                            If PointCompare.DistanceTo(PointBase) > Distance Then
                                Distance = PointCompare.DistanceTo(PointBase)
                            End If
                        End If
                    End If
                Next
            Next
            Return Distance
        End Function

        ''' <summary>
        ''' Retorna o ponto de intersecção dado um ponto de origem, um ângulo e a entidade a ser interceptada
        ''' </summary>
        ''' <param name="Point">Ponto de referência</param>
        ''' <param name="DegreeAngle">Ângulo (Graus)</param>
        ''' <param name="Entity">Entidade ser interceptada</param>
        ''' <param name="FilterEntity">Filtro de entidades</param>
        ''' <param name="SidePoint">Ponto de referência (Se informado retorna apenas o ponto mais próximo)</param>
        ''' <param name="IntersectType">Modo de cálculo para intersecção</param>
        ''' <param name="MaxLimit">Limite máximo entre SidePoint e a intersecção calculada</param>
        ''' <param name="MinLimit">Limite mínimo entre SidePoint e a intersecção calculada</param>
        ''' <param name="SubItems">Determina se os subitens da entidade serão considerados</param>
        ''' <param name="Reach">Alcance para cálculo polar da intersecção</param>
        ''' <returns>Point3dCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function VirtualIntersection(Point As Point3d, DegreeAngle As Double, Entity As Entity, FilterEntity As ArrayList, Optional SidePoint As Object = Nothing, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendThis, Optional MaxLimit As Double = -1, Optional MinLimit As Double = -1, Optional SubItems As Boolean = True, Optional Reach As Double = 1000) As Point3dCollection
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Point2 As Point3d
            Dim Points As New Point3dCollection
            Dim Point3dCollection As New Point3dCollection
            Dim MinPoint As Object = Nothing
            Dim Line As Line
            Dim DBObjectCollection As DBObjectCollection
            DBObjectCollection = Engine2.EntityInteration.GetSubEntitys(Entity, SubItems, , FilterEntity)
            Point2 = Engine2.Geometry.PolarPoint3d(Point, DegreeAngle, Reach)
            Line = New Line(Point, Point2)
            Try
                Points.Clear()
                For Each DBObject As DBObject In DBObjectCollection
                    Try
                        Line.IntersectWith(Engine2.ConvertObject.DBObjectToEntity(DBObject), IntersectType, Points, IntPtr.Zero, IntPtr.Zero)
                        For Each Point3d As Point3d In Points
                            If Point3dCollection.Contains(Point3d) = False Then
                                Point3dCollection.Add(Point3d)
                            End If
                        Next
                    Catch
                        Exit Try
                    End Try
                Next
                If Point3dCollection.Count > 0 Then
                    If IsNothing(SidePoint) = True Then
                        Return Point3dCollection
                    Else
                        If MinLimit <> -1 Then
                            For Each Point3d As Point3d In Point3dCollection
                                If SidePoint.DistanceTo(Point3d) <= MinLimit Then
                                    Point3dCollection.Remove(Point3d)
                                End If
                            Next
                        End If
                        Point = Engine2.Geometry.GetClosestPoint(Point3dCollection, SidePoint)
                        If IsNothing(Point) = False Then
                            If MaxLimit = -1 Then
                                Point3dCollection.Clear()
                                Point3dCollection.Add(Point)
                                Return Point3dCollection
                            Else
                                If Point.DistanceTo(SidePoint) <= MaxLimit Then
                                    Point3dCollection.Clear()
                                    Point3dCollection.Add(Point)
                                    Return Point3dCollection
                                Else
                                    Return New Point3dCollection
                                End If
                            End If
                        Else
                            Return New Point3dCollection
                        End If
                    End If
                Else
                    Return New Point3dCollection
                End If
            Catch
                Return New Point3dCollection
            End Try
        End Function

        ''' <summary>
        ''' Retorna os pontos de intersecção entre 2 segmentos
        ''' </summary>
        ''' <param name="Curve1">Segmento 1</param>
        ''' <param name="Curve2">Segmento 2</param>
        ''' <param name="SidePoint">Ponto de referência (Se informado retorna apenas o ponto mais próximo)</param>
        ''' <param name="IntersectType">Modo de cálculo para intersecção</param>
        ''' <param name="Limit">Limite máximo entre SidePoint e a intersecção calculada</param>
        ''' <returns>Point3dCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function VirtualIntersection(Curve1 As Curve, Curve2 As Curve, Optional SidePoint As Object = Nothing, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendThis, Optional Limit As Double = -1) As Point3dCollection
            Dim Points As New Point3dCollection
            Dim Point3dCollection As New Point3dCollection
            Dim MinPoint As Object = Nothing
            Dim Point As Object
            Try
                Points.Clear()
                Curve1.IntersectWith(Curve2, IntersectType, Points, IntPtr.Zero, IntPtr.Zero)
                For Each Point3d As Point3d In Points
                    If Point3dCollection.Contains(Point3d) = False Then
                        Point3dCollection.Add(Point3d)
                    End If
                Next
                If Point3dCollection.Count > 0 Then
                    If IsNothing(SidePoint) = True Then
                        Return Point3dCollection
                    Else
                        Point = Engine2.Geometry.GetClosestPoint(Point3dCollection, SidePoint)
                        If IsNothing(Point) = False Then
                            If Limit = -1 Then
                                Point3dCollection.Clear()
                                Point3dCollection.Add(Point)
                                Return Point3dCollection
                            Else
                                If Point.DistanceTo(SidePoint) <= Limit Then
                                    Point3dCollection.Clear()
                                    Point3dCollection.Add(Point)
                                    Return Point3dCollection
                                Else
                                    Return New Point3dCollection
                                End If
                            End If
                        Else
                            Return New Point3dCollection
                        End If
                    End If
                Else
                    Return New Point3dCollection
                End If
            Catch
                Return New Point3dCollection
            End Try
        End Function

        ''' <summary>
        ''' Retorna os pontos de intersecção entre 2 segmentos
        ''' </summary>
        ''' <param name="LineSegment3d1">Segmento 1</param>
        ''' <param name="LineSegment3d2">Segmento 2</param>
        ''' <param name="SidePoint">Ponto de referência (Se informado retorna apenas o ponto mais próximo)</param>
        ''' <param name="IntersectType">Modo de cálculo para intersecção</param>
        ''' <param name="Limit">Limite máximo entre SidePoint e a intersecção calculada</param>
        ''' <returns>Point3dCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function VirtualIntersection(LineSegment3d1 As LineSegment3d, LineSegment3d2 As LineSegment3d, Optional SidePoint As Object = Nothing, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendThis, Optional Limit As Double = -1) As Point3dCollection
            Dim Points As New Point3dCollection
            Dim Point3dCollection As New Point3dCollection
            Dim MinPoint As Object = Nothing
            Dim Point As Object
            Dim Curve1 As Curve
            Dim Curve2 As Curve
            Try
                Points.Clear()
                Curve1 = New Line(LineSegment3d1.StartPoint, LineSegment3d1.EndPoint)
                Curve2 = New Line(LineSegment3d2.StartPoint, LineSegment3d2.EndPoint)
                Curve1.IntersectWith(Curve2, IntersectType, Points, IntPtr.Zero, IntPtr.Zero)
                For Each Point3d As Point3d In Points
                    If Point3dCollection.Contains(Point3d) = False Then
                        Point3dCollection.Add(Point3d)
                    End If
                Next
                If Point3dCollection.Count > 0 Then
                    If IsNothing(SidePoint) = True Then
                        Return Point3dCollection
                    Else
                        Point = Engine2.Geometry.GetClosestPoint(Point3dCollection, SidePoint)
                        If IsNothing(Point) = False Then
                            If Limit = -1 Then
                                Point3dCollection.Clear()
                                Point3dCollection.Add(Point)
                                Return Point3dCollection
                            Else
                                If Point.DistanceTo(SidePoint) <= Limit Then
                                    Point3dCollection.Clear()
                                    Point3dCollection.Add(Point)
                                    Return Point3dCollection
                                Else
                                    Return New Point3dCollection
                                End If
                            End If
                        Else
                            Return New Point3dCollection
                        End If
                    End If
                Else
                    Return New Point3dCollection
                End If
            Catch
                Return New Point3dCollection
            End Try
        End Function

        ''' <summary>
        ''' Retorna os pontos de intersecção entre uma coleção de coleções de pontos que compõe segmentos de reta virtuais
        ''' </summary>
        ''' <param name="VirtualParallelPoints">Coleção de pontos que representam as retas a serem avaliadas</param>
        ''' <param name="SidePoint">Ponto de referência (Se informado retorna apenas o ponto mais próximo)</param>
        ''' <param name="IntersectType">Modo de cálculo para intersecção</param>
        ''' <param name="Limit">Limite máximo entre SidePoint e a intersecção calculada</param>
        ''' <returns>Point3dCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function VirtualIntersection(VirtualParallelPoints As List(Of Point3dCollection), Optional SidePoint As Object = Nothing, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendThis, Optional Limit As Double = -1) As Point3dCollection
            Dim Points As New Point3dCollection
            Dim Point3dCollection As New Point3dCollection
            Dim Lines As New List(Of Line)
            Dim MinPoint As Object = Nothing
            Dim Point As Object
            Try
                For Each PointCollection As Point3dCollection In VirtualParallelPoints
                    Lines.Add(New Line(PointCollection.Item(0), PointCollection.Item(1)))
                Next
                For Each line As Line In Lines
                    For Each line2 As Line In Lines
                        If line.Equals(line2) = False Then
                            Points.Clear()
                            line.IntersectWith(line2, IntersectType, Points, IntPtr.Zero, IntPtr.Zero)
                            For Each Point3d As Point3d In Points
                                If Point3dCollection.Contains(Point3d) = False Then
                                    Point3dCollection.Add(Point3d)
                                End If
                            Next
                        End If
                    Next
                Next
                If Point3dCollection.Count > 0 Then
                    If IsNothing(SidePoint) = True Then
                        Return Point3dCollection
                    Else
                        Point = Engine2.Geometry.GetClosestPoint(Point3dCollection, SidePoint)
                        If IsNothing(Point) = False Then
                            If Limit = -1 Then
                                Point3dCollection.Clear()
                                Point3dCollection.Add(Point)
                                Return Point3dCollection
                            Else
                                If Point.DistanceTo(SidePoint) <= Limit Then
                                    Point3dCollection.Clear()
                                    Point3dCollection.Add(Point)
                                    Return Point3dCollection
                                Else
                                    Return New Point3dCollection
                                End If
                            End If
                        Else
                            Return New Point3dCollection
                        End If
                    End If
                Else
                    Return New Point3dCollection
                End If
            Catch
                Return New Point3dCollection
            End Try
        End Function

        ''' <summary>
        ''' Retorna o ponto de intersecção dado um ponto de origem, um ângulo e a entidade a ser interceptada 
        ''' </summary>
        ''' <param name="Point">Ponto de referência</param>
        ''' <param name="DegreeAngle">Ângulo (Graus)</param>
        ''' <param name="LineSegment2d">Entidade a ser interceptado</param>
        ''' <param name="IntersectType">Modo de cálculo para intersecção</param>
        ''' <param name="Reach">Alcance para cálculo polar da intersecção</param>
        ''' <returns>Point3dCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function VirtualIntersection(Point As Point3d, DegreeAngle As Double, LineSegment2d As LineSegment2d, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendThis, Optional Reach As Double = 1000) As Point3dCollection
            Dim Point2 As Point3d
            Dim VirtualParallelPoints As New List(Of Point3dCollection)
            Dim Point3dCollection As Point3dCollection
            Point2 = Engine2.Geometry.PolarPoint3d(Point, DegreeAngle, Reach)
            Point3dCollection = New Point3dCollection
            Point3dCollection.Add(Point)
            Point3dCollection.Add(Point2)
            VirtualParallelPoints.Add(Point3dCollection)
            Point3dCollection = New Point3dCollection
            Point3dCollection.Add(Engine2.Geometry.Point2dToPoint3d(LineSegment2d.StartPoint))
            Point3dCollection.Add(Engine2.Geometry.Point2dToPoint3d(LineSegment2d.EndPoint))
            VirtualParallelPoints.Add(Point3dCollection)
            Return Engine2.Geometry.VirtualIntersection(VirtualParallelPoints, Nothing, IntersectType, -1)
        End Function

        ''' <summary>
        ''' Retorna o ponto de intersecção dado um ponto de origem, um ângulo e a entidade a ser interceptada 
        ''' </summary>
        ''' <param name="Point">Ponto de referência</param>
        ''' <param name="DegreeAngle">Ângulo (Graus)</param>
        ''' <param name="LineSegment3d">Entidade a ser interceptada</param>
        ''' <param name="IntersectType">Modo de cálculo para intersecção</param>
        ''' <param name="Reach">Alcance para cálculo polar da intersecção</param>
        ''' <returns>Point3dCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function VirtualIntersection(Point As Point3d, DegreeAngle As Double, LineSegment3d As LineSegment3d, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendThis, Optional Reach As Double = 1000) As Point3dCollection
            Dim Point2 As Point3d
            Dim VirtualParallelPoints As New List(Of Point3dCollection)
            Dim Point3dCollection As Point3dCollection
            Point2 = Engine2.Geometry.PolarPoint3d(Point, DegreeAngle, Reach)
            Point3dCollection = New Point3dCollection
            Point3dCollection.Add(Point)
            Point3dCollection.Add(Point2)
            VirtualParallelPoints.Add(Point3dCollection)
            Point3dCollection = New Point3dCollection
            Point3dCollection.Add(LineSegment3d.StartPoint)
            Point3dCollection.Add(LineSegment3d.EndPoint)
            VirtualParallelPoints.Add(Point3dCollection)
            Return Engine2.Geometry.VirtualIntersection(VirtualParallelPoints, Nothing, IntersectType, -1)
        End Function


        ''' <summary>
        ''' Retorna o ponto de intersecção dado um ponto de origem, um ângulo e a entidade a ser interceptada 
        ''' </summary>
        ''' <param name="Point">Ponto de referência</param>
        ''' <param name="DegreeAngle">Ângulo (Graus)</param>
        ''' <param name="Line2d">Entidade a ser interceptado</param>
        ''' <param name="IntersectType">Modo de cálculo para intersecção</param>
        ''' <param name="Reach">Alcance para cálculo polar da intersecção</param>
        ''' <returns>Point3dCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function VirtualIntersection(Point As Point3d, DegreeAngle As Double, Line2d As Line2d, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendThis, Optional Reach As Double = 1000) As Point3dCollection
            Dim Point2 As Point3d
            Dim VirtualParallelPoints As New List(Of Point3dCollection)
            Dim Point3dCollection As Point3dCollection
            Point2 = Engine2.Geometry.PolarPoint3d(Point, DegreeAngle, Reach)
            Point3dCollection = New Point3dCollection
            Point3dCollection.Add(Point)
            Point3dCollection.Add(Point2)
            VirtualParallelPoints.Add(Point3dCollection)
            Point3dCollection = New Point3dCollection
            Point3dCollection.Add(Engine2.Geometry.Point2dToPoint3d(Line2d.StartPoint))
            Point3dCollection.Add(Engine2.Geometry.Point2dToPoint3d(Line2d.EndPoint))
            VirtualParallelPoints.Add(Point3dCollection)
            Return Engine2.Geometry.VirtualIntersection(VirtualParallelPoints, Nothing, IntersectType, -1)
        End Function

        ''' <summary>
        ''' Retorna o ponto de intersecção dado um ponto de origem, um ângulo e a entidade a ser interceptada 
        ''' </summary>
        ''' <param name="Point">Ponto de referência</param>
        ''' <param name="DegreeAngle">Ângulo (Graus)</param>
        ''' <param name="Line3d">Entidade a ser interceptada</param>
        ''' <param name="IntersectType">Modo de cálculo para intersecção</param>
        ''' <param name="Reach">Alcance para cálculo polar da intersecção</param>
        ''' <returns>Point3dCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function VirtualIntersection(Point As Point3d, DegreeAngle As Double, Line3d As Line3d, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendThis, Optional Reach As Double = 1000) As Point3dCollection
            Dim Point2 As Point3d
            Dim VirtualParallelPoints As New List(Of Point3dCollection)
            Dim Point3dCollection As Point3dCollection
            Point2 = Engine2.Geometry.PolarPoint3d(Point, DegreeAngle, Reach)
            Point3dCollection = New Point3dCollection
            Point3dCollection.Add(Point)
            Point3dCollection.Add(Point2)
            VirtualParallelPoints.Add(Point3dCollection)
            Point3dCollection = New Point3dCollection
            Point3dCollection.Add(Line3d.StartPoint)
            Point3dCollection.Add(Line3d.EndPoint)
            VirtualParallelPoints.Add(Point3dCollection)
            Return Engine2.Geometry.VirtualIntersection(VirtualParallelPoints, Nothing, IntersectType, -1)
        End Function

        ''' <summary>
        ''' Calcula os pontos paralelos 
        ''' </summary>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="EndPoint">Ponto final</param>
        ''' <param name="OffSet">Offset</param>
        ''' <returns>List(Of Point3dCollection)</returns>
        ''' <remarks></remarks>
        Public Shared Function VirtualParallelPoints(StartPoint As Point3d, EndPoint As Point3d, OffSet As Double) As List(Of Point3dCollection)
            Try
                VirtualParallelPoints = New List(Of Point3dCollection)
                Dim Angle As Double = Engine2.Geometry.GetAngle(StartPoint, EndPoint, eAngleFormat.Degrees)
                Dim ParallelPoints1 As New Point3dCollection
                Dim ParallelPoints2 As New Point3dCollection
                ParallelPoints1.Add(Engine2.Geometry.PolarPoint3d(StartPoint, Angle + 90, OffSet))
                ParallelPoints1.Add(Engine2.Geometry.PolarPoint3d(EndPoint, Angle + 90, OffSet))
                ParallelPoints2.Add(Engine2.Geometry.PolarPoint3d(StartPoint, Angle + 270, OffSet))
                ParallelPoints2.Add(Engine2.Geometry.PolarPoint3d(EndPoint, Angle + 270, OffSet))
                VirtualParallelPoints.Add(ParallelPoints1)
                VirtualParallelPoints.Add(ParallelPoints2)
                Return VirtualParallelPoints
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Calcula os pontos paralelos 
        ''' </summary>
        ''' <param name="LineSegment3d">Segmento de reta</param>
        ''' <param name="OffSet">Offset</param>
        ''' <returns>List(Of Point3dCollection)</returns>
        ''' <remarks></remarks>
        Public Shared Function VirtualParallelPoints(LineSegment3d As LinearEntity3d, OffSet As Double) As List(Of Point3dCollection)
            Return Engine2.Geometry.VirtualParallelPoints(LineSegment3d.StartPoint, LineSegment3d.EndPoint, OffSet)
        End Function

        ''' <summary>
        ''' Calcula o ângulo entre dois segmentos
        ''' </summary>
        ''' <param name="Line1">Segmento 1</param>
        ''' <param name="Line2">Segmento 2</param>
        ''' <param name="AngleFormat">Formato de saída</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetAngle(Line1 As Line, Line2 As Line, Optional AngleFormat As eAngleFormat = eAngleFormat.Radian) As Double
            Dim v1 As Vector3d = Line1.EndPoint - Line1.StartPoint
            Dim v2 As Vector3d = Line2.EndPoint - Line2.StartPoint
            Select Case AngleFormat
                Case eAngleFormat.Radian
                    Return v1.GetAngleTo(v2, Vector3d.ZAxis.Negate)
                Case eAngleFormat.Degrees
                    Return Engine2.Geometry.RadianToDegree(v1.GetAngleTo(v2, Vector3d.ZAxis.Negate), True)
            End Select
        End Function

        ''' <summary>
        ''' Calcula o ângulo entre dois segmentos
        ''' </summary>
        ''' <param name="LineSegment1">Segmento 1</param>
        ''' <param name="LineSegment2">Segmento 2</param>
        ''' <param name="AngleFormat">Formato de saída</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetAngle(LineSegment1 As LineSegment3d, LineSegment2 As LineSegment3d, Optional AngleFormat As eAngleFormat = eAngleFormat.Radian) As Double
            Dim v1 As Vector3d = LineSegment1.EndPoint - LineSegment1.StartPoint
            Dim v2 As Vector3d = LineSegment2.EndPoint - LineSegment2.StartPoint
            Select Case AngleFormat
                Case eAngleFormat.Radian
                    Return v1.GetAngleTo(v2, Vector3d.ZAxis.Negate)
                Case eAngleFormat.Degrees
                    Return Engine2.Geometry.RadianToDegree(v1.GetAngleTo(v2, Vector3d.ZAxis.Negate), True)
            End Select
        End Function

        ''' <summary>
        ''' Calcula o ângulo entre dois segmentos
        ''' </summary>
        ''' <param name="VerticePoint">Ponto do vértice 3d</param>
        ''' <param name="Point1">Ponto 1 3d</param>
        ''' <param name="Point2">Ponto 2 3d</param>
        ''' <param name="AngleFormat">Formato de saída</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetAngle(VerticePoint As Point3d, Point1 As Point3d, Point2 As Point3d, Optional AngleFormat As eAngleFormat = eAngleFormat.Radian) As Double
            Dim LineSegment1 As LineSegment3d = New LineSegment3d(VerticePoint, Point1)
            Dim LineSegment2 As LineSegment3d = New LineSegment3d(VerticePoint, Point2)
            Return Engine2.Geometry.GetAngle(LineSegment1, LineSegment2, AngleFormat)
        End Function

        ''' <summary>
        ''' Calcula o ângulo entre dois segmentos
        ''' </summary>
        ''' <param name="VerticePoint">Ponto do vértice 2d</param>
        ''' <param name="Point1">Ponto 1 2d</param>
        ''' <param name="Point2">Ponto 2 2d</param>
        ''' <param name="AngleFormat">Formato de saída</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetAngle(VerticePoint As Point2d, Point1 As Point2d, Point2 As Point2d, Optional AngleFormat As eAngleFormat = eAngleFormat.Radian) As Double
            Dim LineSegment1 As LineSegment3d = New LineSegment3d(Engine2.Geometry.Point2dToPoint3d(VerticePoint), Engine2.Geometry.Point2dToPoint3d(Point1))
            Dim LineSegment2 As LineSegment3d = New LineSegment3d(Engine2.Geometry.Point2dToPoint3d(VerticePoint), Engine2.Geometry.Point2dToPoint3d(Point2))
            Return Engine2.Geometry.GetAngle(LineSegment1, LineSegment2, AngleFormat)
        End Function

        ''' <summary>
        ''' Calcula ângulo entre dois pontos
        ''' </summary>
        ''' <param name="Point1">Ponto 1</param>
        ''' <param name="Point2">Ponto 2</param>
        ''' <param name="AngleFormat">Formato de saída</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetAngle(Point1 As Point2d, Point2 As Point2d, Optional AngleFormat As eAngleFormat = eAngleFormat.Radian) As Double
            Select Case AngleFormat
                Case eAngleFormat.Radian
                    Return Point1.GetVectorTo(Point2).Angle
                Case eAngleFormat.Degrees
                    Return Engine2.Geometry.RadianToDegree(Point1.GetVectorTo(Point2).Angle, True)
            End Select
        End Function

        ''' <summary>
        ''' Calcula ângulo entre dois pontos
        ''' </summary>
        ''' <param name="Point1"></param>
        ''' <param name="Point2"></param>
        ''' <param name="AngleFormat">Formato de saída</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetAngle(Point1 As Point3d, Point2 As Point3d, Optional AngleFormat As eAngleFormat = eAngleFormat.Radian) As Double
            Select Case AngleFormat
                Case eAngleFormat.Radian
                    Return Engine2.Geometry.Point3dToPoint2d(Point1).GetVectorTo(Engine2.Geometry.Point3dToPoint2d(Point2)).Angle
                Case eAngleFormat.Degrees
                    Return Engine2.Geometry.RadianToDegree(Engine2.Geometry.Point3dToPoint2d(Point1).GetVectorTo(Engine2.Geometry.Point3dToPoint2d(Point2)).Angle, True)
            End Select
        End Function

        ''' <summary>
        ''' Ordena a coleção de pontos 3D
        ''' </summary>
        ''' <param name="Point3dCollection">Coleção de pontos 3D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SortPoint3dCollection(Point3dCollection As Point3dCollection) As Point3dCollection
            Dim Point3dArray As Point3d() = New Point3d(Point3dCollection.Count - 1) {}
            Point3dCollection.CopyTo(Point3dArray, 0)
            Array.Sort(Point3dArray, New Sort3dByX())
            Return New Point3dCollection(Point3dArray)
        End Function

        ''' <summary>
        ''' Ordena a coleção de pontos 2D
        ''' </summary>
        ''' <param name="Point2dCollection">Coleção de pontos 2D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SortPoint2dCollection(Point2dCollection As Point2dCollection) As Point2dCollection
            Dim Point2dArray As Point2d() = Point2dCollection.ToArray()
            Array.Sort(Point2dArray, New Sort2dByX())
            Return New Point2dCollection(Point2dArray)
        End Function

        ''' <summary>
        ''' Retorna os pontos 3D localizados nas duas coleções
        ''' </summary>
        ''' <param name="Collection1">Coleção 1</param>
        ''' <param name="Collection2">Coleção 2</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Point3dCollectionIntersection(Collection1 As Point3dCollection, Collection2 As Point3dCollection) As Point3dCollection
            Point3dCollectionIntersection = New Point3dCollection
            For Each Item1 As Point3d In Collection1
                For Each Item2 As Point3d In Collection2
                    If Item1 = Item2 Then
                        If Point3dCollectionIntersection.Contains(Item1) = False Then
                            Point3dCollectionIntersection.Add(Item1)
                        End If
                    End If
                Next
            Next
            Return Point3dCollectionIntersection
        End Function

        ''' <summary>
        ''' Retorna os pontos 2D localizados nas duas coleções
        ''' </summary>
        ''' <param name="Collection1">Coleção 1</param>
        ''' <param name="Collection2">Coleção 2</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Point2dCollectionIntersection(Collection1 As Point2dCollection, Collection2 As Point2dCollection) As Point2dCollection
            Point2dCollectionIntersection = New Point2dCollection
            For Each Item1 As Point2d In Collection1
                For Each Item2 As Point2d In Collection2
                    If Item1 = Item2 Then
                        If Point2dCollectionIntersection.Contains(Item1) = False Then
                            Point2dCollectionIntersection.Add(Item1)
                        End If
                    End If
                Next
            Next
            Return Point2dCollectionIntersection
        End Function

        ''' <summary>
        ''' Retorna os pontos 3D localizados nas duas coleções
        ''' </summary>
        ''' <param name="Collection1">Coleção 1</param>
        ''' <param name="Collection2">Coleção 2</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ListOfPoint3dCollectionIntersection(Collection1 As List(Of Point3d), Collection2 As List(Of Point3d)) As List(Of Point3d)
            ListOfPoint3dCollectionIntersection = New List(Of Point3d)
            For Each Item1 As Point3d In Collection1
                For Each Item2 As Point3d In Collection2
                    If Item1 = Item2 Then
                        If ListOfPoint3dCollectionIntersection.Contains(Item1) = False Then
                            ListOfPoint3dCollectionIntersection.Add(Item1)
                        End If
                    End If
                Next
            Next
            Return ListOfPoint3dCollectionIntersection
        End Function

        ''' <summary>
        ''' Retorna os pontos 2D localizados nas duas coleções
        ''' </summary>
        ''' <param name="Collection1">Coleção 1</param>
        ''' <param name="Collection2">Coleção 2</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ListOfPoint2dCollectionIntersection(Collection1 As List(Of Point2d), Collection2 As List(Of Point2d)) As List(Of Point2d)
            ListOfPoint2dCollectionIntersection = New List(Of Point2d)
            For Each Item1 As Point2d In Collection1
                For Each Item2 As Point2d In Collection2
                    If Item1 = Item2 Then
                        If ListOfPoint2dCollectionIntersection.Contains(Item1) = False Then
                            ListOfPoint2dCollectionIntersection.Add(Item1)
                        End If
                    End If
                Next
            Next
            Return ListOfPoint2dCollectionIntersection
        End Function

        ''' <summary>
        ''' Converte a coleção de pontos 3D para Point3dCollection
        ''' </summary>
        ''' <param name="ListOfPoint3d">Coleção de pontos 3D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ListOfPoint3dToPoint3dCollection(ListOfPoint3d As List(Of Point3d)) As Object
            Return New Point3dCollection(ListOfPoint3d.ToArray)
        End Function

        ''' <summary>
        ''' Converte a coleção de pontos 2D para Point2dCollection
        ''' </summary>
        ''' <param name="ListOfPoint2d">Coleção de pontos 2D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ListOfPoint2dToPoint2dCollection(ListOfPoint2d As List(Of Point2d)) As Object
            Return New Point2dCollection(ListOfPoint2d.ToArray)
        End Function

        ''' <summary>
        ''' Obtem os pontos referente aos vértices de uma entidade gráfica dos tipos POLYLINE, POLYLINE2D, POLYLINE3D, LINE
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="IncludeMidpointArc">Determina se em arcos será incluído o ponto médio</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetPoints3d(Entity As Entity, Optional IncludeMidpointArc As Boolean = False) As List(Of Point3d)
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim DBObject As DBObject
                Dim Polyline As Polyline
                Dim Polyline2d As Polyline2d
                Dim Polyline3d As Polyline3d
                Dim Vertex As Integer
                Dim Vertex2d As Vertex2d
                Dim Vertex3d As PolylineVertex3d
                Dim Arc As Arc
                Dim Ellipse As Ellipse
                Dim Line As Line
                Dim Circle As Circle
                Dim Region As Region
                Dim Points = New List(Of Point3d)
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        DBObject = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
                        Select Case Entity.GetType.Name
                            Case "Polyline"
                                Polyline = TryCast(DBObject, Polyline)
                                Vertex = Polyline.NumberOfVertices
                                For VertexIndex As Integer = 0 To Vertex - 1
                                    Points.Add(Polyline.GetPoint3dAt(VertexIndex))
                                Next
                            Case "Polyline2d"
                                Polyline2d = TryCast(DBObject, Polyline2d)
                                For Each VertexID As ObjectId In Polyline2d
                                    Vertex2d = DirectCast(Transaction.GetObject(VertexID, OpenMode.ForRead), Vertex2d)
                                    Points.Add(New Point3d(Vertex2d.Position.X, Vertex2d.Position.Y, 0))
                                Next
                            Case "Polyline3d"
                                Polyline3d = TryCast(DBObject, Polyline3d)
                                For Each VertexID As ObjectId In Polyline3d
                                    Vertex3d = DirectCast(Transaction.GetObject(VertexID, OpenMode.ForRead), PolylineVertex3d)
                                    Points.Add(Vertex3d.Position)
                                Next
                            Case "Line"
                                Line = TryCast(DBObject, Line)
                                Points.Add(Line.StartPoint)
                                Points.Add(Line.EndPoint)
                            Case "Arc"
                                Arc = TryCast(DBObject, Arc)
                                Points.Add(Arc.StartPoint)
                                If IncludeMidpointArc = True Then
                                    Points.Add(Engine2.Geometry.MidPoint(Arc))
                                End If
                                Points.Add(Arc.EndPoint)
                            Case "Ellipse"
                                Ellipse = TryCast(DBObject, Ellipse)
                                Points.Add(Ellipse.StartPoint)
                                Points.Add(Ellipse.EndPoint)
                            Case "Circle"
                                Circle = TryCast(DBObject, Circle)
                                Points.Add(Circle.StartPoint)
                                Points.Add(Circle.EndPoint)
                            Case "Region"
                                Region = TryCast(DBObject, Region)
                                Points.Add(Region.GeometricExtents.MaxPoint())
                                Points.Add(Region.GeometricExtents.MinPoint())
                            Case Else
                                Throw New System.Exception("O tipo " & Entity.GetType.Name & " não é válido para Geometry.GetPoints")
                        End Select
                        Transaction.Commit()
                    End Using
                End Using
                Return Points
            Catch ex As System.Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem os pontos referente aos vértices de uma entidade gráfica dos tipos POLYLINE, POLYLINE2D, POLYLINE3D, LINE
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="IncludeMidpointArc">Determina se em arcos será incluído o ponto médio</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetPoints2d(Entity As Entity, Optional IncludeMidpointArc As Boolean = False) As List(Of Point2d)
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim DBObject As DBObject
                Dim Polyline As Polyline
                Dim Polyline2d As Polyline2d
                Dim Polyline3d As Polyline3d
                Dim Vertex As Integer
                Dim Vertex2d As Vertex2d
                Dim Vertex3d As PolylineVertex3d
                Dim Arc As Arc
                Dim Ellipse As Ellipse
                Dim Line As Line
                Dim Circle As Circle
                Dim Region As Region
                Dim Points = New List(Of Point2d)
                Using DocumentLock As DocumentLock = Document.LockDocument()
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        DBObject = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
                        Select Case Entity.GetType.Name
                            Case "Polyline"
                                Polyline = TryCast(DBObject, Polyline)
                                Vertex = Polyline.NumberOfVertices
                                For VertexIndex As Integer = 0 To Vertex - 1
                                    Points.Add(Polyline.GetPoint2dAt(VertexIndex))
                                Next
                            Case "Polyline2d"
                                Polyline2d = TryCast(DBObject, Polyline2d)
                                For Each VertexID As ObjectId In Polyline2d
                                    Vertex2d = DirectCast(Transaction.GetObject(VertexID, OpenMode.ForRead), Vertex2d)
                                    Points.Add(New Point2d(Vertex2d.Position.X, Vertex2d.Position.Y))
                                Next
                            Case "Polyline3d"
                                Polyline3d = TryCast(DBObject, Polyline3d)
                                For Each VertexID As ObjectId In Polyline3d
                                    Vertex3d = DirectCast(Transaction.GetObject(VertexID, OpenMode.ForRead), PolylineVertex3d)
                                    Points.Add(New Point2d(Vertex3d.Position.X, Vertex3d.Position.Y))
                                Next
                            Case "Line"
                                Line = TryCast(DBObject, Line)
                                Points.Add(New Point2d(Line.StartPoint.X, Line.StartPoint.Y))
                                Points.Add(New Point2d(Line.EndPoint.X, Line.EndPoint.Y))
                            Case "Arc"
                                Arc = TryCast(DBObject, Arc)
                                Points.Add(New Point2d(Arc.StartPoint.X, Arc.StartPoint.Y))
                                If IncludeMidpointArc = True Then
                                    Points.Add(Engine2.Geometry.Point3dToPoint2d(Engine2.Geometry.MidPoint(Arc)))
                                End If
                                Points.Add(New Point2d(Arc.EndPoint.X, Arc.EndPoint.Y))
                            Case "Ellipse"
                                Ellipse = TryCast(DBObject, Ellipse)
                                Points.Add(New Point2d(Ellipse.StartPoint.X, Ellipse.StartPoint.Y))
                                Points.Add(New Point2d(Ellipse.EndPoint.X, Ellipse.EndPoint.Y))
                            Case "Circle"
                                Circle = TryCast(DBObject, Circle)
                                Points.Add(New Point2d(Circle.StartPoint.X, Circle.StartPoint.Y))
                                Points.Add(New Point2d(Circle.EndPoint.X, Circle.EndPoint.Y))
                            Case "Region"
                                Region = TryCast(DBObject, Region)
                                Points.Add(New Point2d(Region.GeometricExtents.MaxPoint.X, Region.GeometricExtents.MaxPoint.Y))
                                Points.Add(New Point2d(Region.GeometricExtents.MinPoint.X, Region.GeometricExtents.MinPoint.Y))
                            Case Else
                                Throw New System.Exception("O tipo " & Entity.GetType.Name & " não é válido para Geometry.GetPoints")
                        End Select
                        Transaction.Commit()
                    End Using
                End Using
                Return Points
            Catch ex As System.Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem o comprimento (Pode ser utilizado para perímetro em entidades fechadas)
        ''' </summary>
        ''' <param name="Entity">Entity</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetLength(Entity As Entity) As Double
            Try
                Select Case Entity.GetType.Name.ToUpper
                    Case "CIRCLE"
                        Dim Circle As Circle = Entity
                        Return Circle.Circumference
                    Case "REGION"
                        Dim Region As Autodesk.AutoCAD.DatabaseServices.Region = Entity
                        Return Region.Perimeter
                    Case "POLYLINE3D"
                        Dim Polyline3D As Polyline3d = Entity
                        Return Polyline3D.Length
                    Case "POLYLINE"
                        Dim Polyline As Polyline = Entity
                        Return Polyline.Length
                    Case "LINE"
                        Dim Line As Line = Entity
                        Return Line.Length
                    Case "ARC"
                        Dim Arc As Arc = Entity
                        Return Arc.Length
                    Case "ELLIPSE"
                        Dim Ellipse As Ellipse = Entity
                        Return Ellipse.GetDistanceAtParameter(Ellipse.EndParam) - Ellipse.GetDistanceAtParameter(Ellipse.StartParam)
                    Case "SPLINE"
                        Dim Spline As Spline = Entity
                        Return Spline.GetDistanceAtParameter(Spline.EndParam) - Spline.GetDistanceAtParameter(Spline.StartParam)
                    Case Else
                        Throw New System.Exception("O tipo " & Entity.GetType.Name & " não é válido para Geometry.GetLength")
                End Select
            Catch ex As System.Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem o comprimento (Pode ser utilizado para perímetro em entidades fechadas)
        ''' </summary>
        ''' <param name="Curve">Curve</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetLength(Curve As Curve) As Double
            Return Math.Abs(Curve.GetDistanceAtParameter(Curve.GetParameterAtPoint(Curve.StartPoint)) - Curve.GetDistanceAtParameter(Curve.GetParameterAtPoint(Curve.EndPoint)))
        End Function

        ''' <summary>
        ''' Obtem a área da entidade
        ''' </summary>
        ''' <param name="Entity">Entity</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetArea(Entity As Entity) As Double
            Try
                Select Case Entity.GetType.Name.ToUpper
                    Case "POLYLINE"
                        Dim Polyline As Polyline = Entity
                        Return Polyline.Area
                    Case "POLYLINE3D"
                        Dim Polyline3D As Polyline3d = Entity
                        Return Polyline3D.Area
                    Case "CIRCLE"
                        Dim Circle As Circle = Entity
                        Return Circle.Area
                    Case "REGION"
                        Dim Region As Autodesk.AutoCAD.DatabaseServices.Region = Entity
                        Return Region.Area
                    Case "ELLIPSE"
                        Dim Ellipse As Ellipse = Entity
                        Return Ellipse.Area
                    Case "SPLINE"
                        Dim Spline As Spline = Entity
                        Return Spline.Area
                    Case Else
                        Throw New System.Exception("O tipo " & Entity.GetType.Name & " não é válido para Geometry.GetArea")
                End Select
            Catch ex As System.Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retorna o ponto médio de uma entidade (Linha ou Arco)
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MidPoint(Entity As Entity) As Point3d
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim Arc As Arc
                Dim Polyline As Polyline
                Dim Line As Line
                Dim Spline As Spline
                Select Case Entity.GetType.Name
                    Case "Arc"
                        Arc = Entity
                        MidPoint = Arc.GetPointAtDist(Arc.Length / 2)
                    Case "Line"
                        Line = Entity
                        MidPoint = Line.GetPointAtDist(Line.Length / 2)
                    Case "Polyline"
                        Polyline = Entity
                        MidPoint = Polyline.GetPointAtDist(Polyline.Length / 2)
                    Case "Spline"
                        Spline = Entity
                        MidPoint = Spline.GetPointAtDist(Engine2.Geometry.GetLength(Spline) / 2)
                    Case Else
                        Throw New System.Exception("O tipo " & Entity.GetType.Name & " não é válido para Geometry.ArcMidPoint")
                End Select
                Return MidPoint
            Catch ex As System.Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retorna o ponto médio entre 2 pontos
        ''' </summary>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="EndPoint">Ponto final</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MidPoint(StartPoint As Point3d, EndPoint As Point3d) As Point3d
            Return New LineSegment3d(StartPoint, EndPoint).MidPoint
        End Function

        ''' <summary>
        ''' Converte ponto 2d para ponto 3d
        ''' </summary>
        ''' <param name="Point2d"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Point2dToPoint3d(Point2d As Point2d) As Point3d
            Return New Point3d(Point2d.X, Point2d.Y, 0)
        End Function

        ''' <summary>
        ''' Converte ponto 3d para ponto 2d
        ''' </summary>
        ''' <param name="Point3d"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Point3dToPoint2d(Point3d As Point3d) As Point2d
            Return Point3d.Convert2d(New Plane(Point3d.Origin, Vector3d.ZAxis))
        End Function

        ''' <summary>
        ''' Converte ponto 2d para ponto
        ''' </summary>
        ''' <param name="Point2d"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Point2dToPoint(Point2d As Point2d) As System.Drawing.Point
            Return New System.Drawing.Point(Point2d.X, Point2d.Y)
        End Function

        ''' <summary>
        ''' Converte ponto 3d para ponto
        ''' </summary>
        ''' <param name="Point3d"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Point3dToPoint(Point3d As Point3d) As System.Drawing.Point
            Return New System.Drawing.Point(Point3d.X, Point3d.Y)
        End Function

        ''' <summary>
        ''' Forma o ponto mais distante em X, Y, Z do plano WCS a partir de uma lista de pontos 3D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 3D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MaxPoint3d(Points As Point3dCollection) As Point3d
            Dim X As New ArrayList
            Dim Y As New ArrayList
            Dim Z As New ArrayList
            For Each Point3d As Point3d In Points
                If X.Contains(Point3d.X) = False Then
                    X.Add(Point3d.X)
                End If
                If Y.Contains(Point3d.Y) = False Then
                    Y.Add(Point3d.Y)
                End If
                If Z.Contains(Point3d.Z) = False Then
                    Z.Add(Point3d.Z)
                End If
            Next
            X.Sort()
            Y.Sort()
            Z.Sort()
            Return New Point3d(X.Item(X.Count - 1), Y.Item(Y.Count - 1), Z.Item(Z.Count - 1))
        End Function

        ''' <summary>
        ''' Forma o ponto mais distante em X, Y, Z do plano WCS a partir de uma lista de pontos 3D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 3D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MaxPoint3d(Points As List(Of Point3d)) As Point3d
            Dim X As New ArrayList
            Dim Y As New ArrayList
            Dim Z As New ArrayList
            For Each Point3d As Point3d In Points
                If X.Contains(Point3d.X) = False Then
                    X.Add(Point3d.X)
                End If
                If Y.Contains(Point3d.Y) = False Then
                    Y.Add(Point3d.Y)
                End If
                If Z.Contains(Point3d.Z) = False Then
                    Z.Add(Point3d.Z)
                End If
            Next
            X.Sort()
            Y.Sort()
            Z.Sort()
            Return New Point3d(X.Item(X.Count - 1), Y.Item(Y.Count - 1), Z.Item(Z.Count - 1))
        End Function

        ''' <summary>
        ''' Forma o ponto mais distante em X, Y do plano WCS a partir de uma lista de pontos 2D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 2D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MaxPoint2d(Points As Point2dCollection) As Point2d
            Dim X As New ArrayList
            Dim Y As New ArrayList
            For Each Point2d As Point2d In Points
                If X.Contains(Point2d.X) = False Then
                    X.Add(Point2d.X)
                End If
                If Y.Contains(Point2d.Y) = False Then
                    Y.Add(Point2d.Y)
                End If
            Next
            X.Sort()
            Y.Sort()
            Return New Point2d(X.Item(X.Count - 1), Y.Item(Y.Count - 1))
        End Function

        ''' <summary>
        ''' Forma o ponto mais distante em X, Y do plano WCS a partir de uma lista de pontos 2D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 2D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MaxPoint2d(Points As List(Of Point2d)) As Point2d
            Dim X As New ArrayList
            Dim Y As New ArrayList
            For Each Point2d As Point2d In Points
                If X.Contains(Point2d.X) = False Then
                    X.Add(Point2d.X)
                End If
                If Y.Contains(Point2d.Y) = False Then
                    Y.Add(Point2d.Y)
                End If
            Next
            X.Sort()
            Y.Sort()
            Return New Point2d(X.Item(X.Count - 1), Y.Item(Y.Count - 1))
        End Function

        ''' <summary>
        ''' Forma o ponto mais distante em X, Y do plano WCS a partir de uma lista de pontos
        ''' </summary>
        ''' <param name="Points">Lista de pontos</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MaxPoint(Points As List(Of System.Drawing.Point)) As System.Drawing.Point
            Dim X As New ArrayList
            Dim Y As New ArrayList
            For Each Point As System.Drawing.Point In Points
                If X.Contains(Point.X) = False Then
                    X.Add(Point.X)
                End If
                If Y.Contains(Point.Y) = False Then
                    Y.Add(Point.Y)
                End If
            Next
            X.Sort()
            Y.Sort()
            Return New System.Drawing.Point(X.Item(X.Count - 1), Y.Item(Y.Count - 1))
        End Function

        ''' <summary>
        ''' Forma o ponto menos distante em X, Y, Z do plano WCS a partir de uma lista de pontos 3D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 3D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MinPoint3d(Points As Point3dCollection) As Point3d
            Dim X As New ArrayList
            Dim Y As New ArrayList
            Dim Z As New ArrayList
            For Each Point3d As Point3d In Points
                If X.Contains(Point3d.X) = False Then
                    X.Add(Point3d.X)
                End If
                If Y.Contains(Point3d.Y) = False Then
                    Y.Add(Point3d.Y)
                End If
                If Z.Contains(Point3d.Z) = False Then
                    Z.Add(Point3d.Z)
                End If
            Next
            X.Sort()
            Y.Sort()
            Z.Sort()
            Return New Point3d(X.Item(0), Y.Item(0), Z.Item(0))
        End Function

        ''' <summary>
        ''' Forma o ponto menos distante em X, Y, Z do plano WCS a partir de uma lista de pontos 3D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 3D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MinPoint3d(Points As List(Of Point3d)) As Point3d
            Dim X As New ArrayList
            Dim Y As New ArrayList
            Dim Z As New ArrayList
            For Each Point3d As Point3d In Points
                If X.Contains(Point3d.X) = False Then
                    X.Add(Point3d.X)
                End If
                If Y.Contains(Point3d.Y) = False Then
                    Y.Add(Point3d.Y)
                End If
                If Z.Contains(Point3d.Z) = False Then
                    Z.Add(Point3d.Z)
                End If
            Next
            X.Sort()
            Y.Sort()
            Z.Sort()
            Return New Point3d(X.Item(0), Y.Item(0), Z.Item(0))
        End Function

        ''' <summary>
        ''' Forma o ponto menos distante em X, Y do plano WCS a partir de uma lista de pontos 2D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 2D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MinPoint2d(Points As Point2dCollection) As Point2d
            Dim X As New ArrayList
            Dim Y As New ArrayList
            For Each Point2d As Point2d In Points
                If X.Contains(Point2d.X) = False Then
                    X.Add(Point2d.X)
                End If
                If Y.Contains(Point2d.Y) = False Then
                    Y.Add(Point2d.Y)
                End If
            Next
            X.Sort()
            Y.Sort()
            Return New Point2d(X.Item(0), Y.Item(0))
        End Function

        ''' <summary>
        ''' Forma o ponto menos distante em X, Y do plano WCS a partir de uma lista de pontos 2D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 2D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MinPoint2d(Points As List(Of Point2d)) As Point2d
            Dim X As New ArrayList
            Dim Y As New ArrayList
            For Each Point2d As Point2d In Points
                If X.Contains(Point2d.X) = False Then
                    X.Add(Point2d.X)
                End If
                If Y.Contains(Point2d.Y) = False Then
                    Y.Add(Point2d.Y)
                End If
            Next
            X.Sort()
            Y.Sort()
            Return New Point2d(X.Item(0), Y.Item(0))
        End Function

        ''' <summary>
        ''' Forma o ponto menos distante em X, Y do plano WCS a partir de uma lista de pontos 2D
        ''' </summary>
        ''' <param name="Points">Lista de pontos</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MinPoint(Points As List(Of System.Drawing.Point)) As System.Drawing.Point
            Dim X As New ArrayList
            Dim Y As New ArrayList
            For Each Point As System.Drawing.Point In Points
                If X.Contains(Point.X) = False Then
                    X.Add(Point.X)
                End If
                If Y.Contains(Point.Y) = False Then
                    Y.Add(Point.Y)
                End If
            Next
            X.Sort()
            Y.Sort()
            Return New System.Drawing.Point(X.Item(0), Y.Item(0))
        End Function

        ''' <summary>
        ''' Forma o ponto médio dado a lista de pontos 3D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 3D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CenterPoint(Points As Point3dCollection) As Point3d
            Dim MaxPoint As Point3d = Engine2.Geometry.MaxPoint3d(Points)
            Dim MinPoint As Point3d = Engine2.Geometry.MinPoint3d(Points)
            Return New Point3d(((MaxPoint.X + MinPoint.X) / 2), ((MaxPoint.Y + MinPoint.Y) / 2), ((MaxPoint.Z + MinPoint.Z) / 2))
        End Function

        ''' <summary>
        ''' Forma o ponto médio dado a lista de pontos 3D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 3D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CenterPoint(Points As List(Of Point3d)) As Point3d
            Dim MaxPoint As Point3d = Engine2.Geometry.MaxPoint3d(Points)
            Dim MinPoint As Point3d = Engine2.Geometry.MinPoint3d(Points)
            Return New Point3d(((MaxPoint.X + MinPoint.X) / 2), ((MaxPoint.Y + MinPoint.Y) / 2), ((MaxPoint.Z + MinPoint.Z) / 2))
        End Function

        ''' <summary>
        ''' Forma o ponto médio dado a lista de pontos 2D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 2D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CenterPoint(Points As Point2dCollection) As Point2d
            Dim MaxPoint As Point2d = Engine2.Geometry.MaxPoint2d(Points)
            Dim MinPoint As Point2d = Engine2.Geometry.MinPoint2d(Points)
            Return New Point2d(((MaxPoint.X + MinPoint.X) / 2), ((MaxPoint.Y + MinPoint.Y) / 2))
        End Function

        ''' <summary>
        ''' Forma o ponto médio dado a lista de pontos 2D
        ''' </summary>
        ''' <param name="Points">Lista de pontos 2D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CenterPoint(Points As List(Of Point2d)) As Point2d
            Dim MaxPoint As Point2d = Engine2.Geometry.MaxPoint2d(Points)
            Dim MinPoint As Point2d = Engine2.Geometry.MinPoint2d(Points)
            Return New Point2d(((MaxPoint.X + MinPoint.X) / 2), ((MaxPoint.Y + MinPoint.Y) / 2))
        End Function

        ''' <summary>
        ''' Forma o ponto médio dado a lista de pontos
        ''' </summary>
        ''' <param name="Points">Lista de pontos</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CenterPoint(Points As List(Of System.Drawing.Point)) As System.Drawing.Point
            Dim MaxPoint As System.Drawing.Point = Engine2.Geometry.MaxPoint(Points)
            Dim MinPoint As System.Drawing.Point = Engine2.Geometry.MinPoint(Points)
            Return New System.Drawing.Point(((MaxPoint.X + MinPoint.X) / 2), ((MaxPoint.Y + MinPoint.Y) / 2))
        End Function

        ''' <summary>
        ''' Obtem o ponto 2d com base em um ângulo, distância e ponto inicial
        ''' </summary>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="DegreeAngle">Ângulo</param>
        ''' <param name="Distance">Distância</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PolarPoint2d(StartPoint As Point2d, DegreeAngle As Double, Distance As Double) As Point2d
            DegreeAngle = DegreeAngle + Math.Ceiling(-DegreeAngle / 360) * 360
            Return New Point2d(Distance * Math.Cos(DegreeAngle * (Math.PI / 180)) + StartPoint.X, Distance * Math.Sin(DegreeAngle * (Math.PI / 180)) + StartPoint.Y)
        End Function

        ''' <summary>
        ''' Obtem o ponto 3d com base em um ângulo, distância e ponto inicial
        ''' </summary>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="DegreeAngle">Ângulo</param>
        ''' <param name="Distance">Distância</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PolarPoint3d(StartPoint As Point3d, DegreeAngle As Double, Distance As Double) As Point3d
            DegreeAngle = DegreeAngle + Math.Ceiling(-DegreeAngle / 360) * 360
            Return New Point3d(Distance * Math.Cos(DegreeAngle * (Math.PI / 180)) + StartPoint.X, Distance * Math.Sin(DegreeAngle * (Math.PI / 180)) + StartPoint.Y, 0)
        End Function

        ''' <summary>
        ''' Converte Graus em Radianos
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DegreeToRadian(Value As Double) As Double
            Return (Value * (Math.PI / 180))
        End Function

        ''' <summary>
        ''' Converte Radianos em Graus
        ''' </summary>
        ''' <param name="Value">Ângulo radiano</param>
        ''' <param name="AngleAjust">Determina se os valores negativos serão reajustados entre 0° e 359°</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function RadianToDegree(Value As Double, Optional AngleAjust As Boolean = False) As Double
            If AngleAjust = True Then
                Return Engine2.Geometry.AngleAjust(Value * (180 / Math.PI))
            Else
                Return (Value * (180 / Math.PI))
            End If
        End Function

        ''' <summary>
        ''' Adição e subtração de ângulos (°)
        ''' </summary>
        ''' <param name="DegreeAngle">Ângulo em graus</param>
        ''' <param name="Value">Value</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AngleAdd(DegreeAngle As Double, Value As Integer) As Double
            DegreeAngle = DegreeAngle + Value
            Return Engine2.Geometry.AngleAjust(DegreeAngle)
        End Function

        ''' <summary>
        ''' Ajusta o valor ângular para que fique dentro de 0° e 359°
        ''' </summary>
        ''' <param name="DegreeAngle">Ângulo em graus</param>
        ''' <param name="TruncateAjust">Determina se o valor considerado será arredondado</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AngleAjust(DegreeAngle As Double, Optional TruncateAjust As Boolean = False) As Double
            If TruncateAjust = True Then
                Return CInt(DegreeAngle) + Math.Ceiling(-CInt(DegreeAngle) / 360) * 360
            Else
                Return DegreeAngle + Math.Ceiling(-DegreeAngle / 360) * 360
            End If
        End Function

        ''' <summary>
        ''' Distância
        ''' </summary>
        ''' <param name="P1">Ponto inicial</param>
        ''' <param name="p2">Ponto final</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Distance(P1 As System.Drawing.Point, p2 As System.Drawing.Point) As Double
            Return Math.Sqrt((Math.Abs(P1.X - p2.X) ^ 2) + (Math.Abs(P1.Y - p2.Y) ^ 2))
        End Function

        ''' <summary>
        ''' Distância
        ''' </summary>
        ''' <param name="P1">Ponto inicial</param>
        ''' <param name="p2">Ponto final</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Distance(P1 As Point2d, p2 As Point2d) As Double
            Return P1.GetDistanceTo(p2)
        End Function

        ''' <summary>
        ''' Distância
        ''' </summary>
        ''' <param name="P1">Ponto inicial</param>
        ''' <param name="p2">Ponto final</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Distance(P1 As Point3d, p2 As Point3d) As Double
            Return P1.DistanceTo(p2)
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo entre o ponto 1 e 2
        ''' </summary>
        ''' <param name="P1">Ponto inicial</param>
        ''' <param name="P2">Ponto final</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSize(P1 As System.Drawing.Point, P2 As System.Drawing.Point) As System.Drawing.Size
            Return New System.Drawing.Size(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo entre o ponto 1 e 2
        ''' </summary>
        ''' <param name="P1">Ponto inicial</param>
        ''' <param name="P2">Ponto final</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSize(P1 As Point2d, P2 As Point2d) As System.Drawing.Size
            Return New System.Drawing.Size(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo entre o ponto 1 e 2
        ''' </summary>
        ''' <param name="P1">Ponto inicial</param>
        ''' <param name="P2">Ponto final</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSize(P1 As Point3d, P2 As Point3d) As System.Drawing.Size
            Return New System.Drawing.Size(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo que abrange todos os pontos
        ''' </summary>
        ''' <param name="Points">Relação de pontos</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSize(Points As List(Of System.Drawing.Point)) As System.Drawing.Size
            Dim P1 As System.Drawing.Point = Engine2.Geometry.MinPoint(Points)
            Dim P2 As System.Drawing.Point = Engine2.Geometry.MaxPoint(Points)
            Return New System.Drawing.Size(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo que abrange todos os pontos
        ''' </summary>
        ''' <param name="Points">Relação de pontos 2D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSize(Points As List(Of Point2d)) As System.Drawing.Size
            Dim P1 As Point2d = Engine2.Geometry.MinPoint2d(Points)
            Dim P2 As Point2d = Engine2.Geometry.MaxPoint2d(Points)
            Return New System.Drawing.Size(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo que abrange todos os pontos
        ''' </summary>
        ''' <param name="Points">Relação de pontos 3D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSize(Points As List(Of Point3d)) As System.Drawing.Size
            Dim P1 As Point3d = Engine2.Geometry.MinPoint3d(Points)
            Dim P2 As Point3d = Engine2.Geometry.MaxPoint3d(Points)
            Return New System.Drawing.Size(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo entre o ponto 1 e 2
        ''' </summary>
        ''' <param name="P1">Ponto inicial</param>
        ''' <param name="P2">Ponto final</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSizef(P1 As System.Drawing.Point, P2 As System.Drawing.Point) As System.Drawing.SizeF
            Return New System.Drawing.SizeF(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo entre o ponto 1 e 2
        ''' </summary>
        ''' <param name="P1">Ponto inicial</param>
        ''' <param name="P2">Ponto final</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSizef(P1 As Point2d, P2 As Point2d) As System.Drawing.SizeF
            Return New System.Drawing.SizeF(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo entre o ponto 1 e 2
        ''' </summary>
        ''' <param name="P1">Ponto inicial</param>
        ''' <param name="P2">Ponto final</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSizef(P1 As Point3d, P2 As Point3d) As System.Drawing.SizeF
            Return New System.Drawing.SizeF(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo que abrange todos os pontos
        ''' </summary>
        ''' <param name="Points">Relação de pontos</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSizef(Points As List(Of System.Drawing.Point)) As System.Drawing.SizeF
            Dim P1 As System.Drawing.Point = Engine2.Geometry.MinPoint(Points)
            Dim P2 As System.Drawing.Point = Engine2.Geometry.MaxPoint(Points)
            Return New System.Drawing.SizeF(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo que abrange todos os pontos
        ''' </summary>
        ''' <param name="Points">Relação de pontos 2D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSizef(Points As List(Of Point2d)) As System.Drawing.SizeF
            Dim P1 As Point2d = Engine2.Geometry.MinPoint2d(Points)
            Dim P2 As Point2d = Engine2.Geometry.MaxPoint2d(Points)
            Return New System.Drawing.SizeF(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Obtem o tamanho considerando um retângulo que abrange todos os pontos
        ''' </summary>
        ''' <param name="Points">Relação de pontos 3D</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetSizef(Points As List(Of Point3d)) As System.Drawing.SizeF
            Dim P1 As Point3d = Engine2.Geometry.MinPoint3d(Points)
            Dim P2 As Point3d = Engine2.Geometry.MaxPoint3d(Points)
            Return New System.Drawing.SizeF(System.Math.Max(P1.X, P2.X) - System.Math.Min(P1.X, P2.X), System.Math.Max(P1.Y, P2.Y) - System.Math.Min(P1.Y, P2.Y))
        End Function

        ''' <summary>
        ''' Calcula o ângulo espelho
        ''' </summary>
        ''' <param name="Angle">Ângulo</param>
        ''' <param name="MidAngle">Ângulo médio para referência</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MirrorAngle(Angle As Double, MidAngle As Double) As Double
            If Angle = 0 And MidAngle = 90 Or MidAngle = 270 Then
                Return 180
            ElseIf Angle = 180 And MidAngle = 90 Or MidAngle = 270 Then
                Return 0
            ElseIf Angle = 90 And MidAngle = 0 Or MidAngle = 180 Then
                Return 270
            ElseIf Angle = 270 And MidAngle = 0 Or MidAngle = 180 Then
                Return 90
            ElseIf MidAngle < Angle Then
                Return Engine2.Geometry.AngleAjust(MidAngle - (Angle - MidAngle))
            ElseIf MidAngle > Angle Then
                Return Engine2.Geometry.AngleAjust(MidAngle + (MidAngle - Angle))
            ElseIf MidAngle = Angle Then
                Return Angle
            End If
        End Function

        ''' <summary>
        ''' Retorna todos os pontos de uma janela retângular dado 2 pontos opostos
        ''' </summary>
        ''' <param name="FirstPoint">Primeiro ponto</param>
        ''' <param name="SecoundPoint">Segundo ponto (oposto)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetWindow(FirstPoint As System.Drawing.Point, SecoundPoint As System.Drawing.Point) As List(Of System.Drawing.Point)
            Dim P1 As System.Drawing.Point
            Dim P2 As System.Drawing.Point
            Dim P3 As System.Drawing.Point
            Dim P4 As System.Drawing.Point
            Dim Points As New List(Of System.Drawing.Point)
            Points.Add(FirstPoint)
            Points.Add(SecoundPoint)
            P1 = Engine2.Geometry.MinPoint(Points)
            P3 = Engine2.Geometry.MaxPoint(Points)
            P2 = New System.Drawing.Point(P3.X, P1.Y)
            P4 = New System.Drawing.Point(P1.X, P3.Y)
            Points.Clear()
            Points.Add(P1)
            Points.Add(P2)
            Points.Add(P3)
            Points.Add(P4)
            Return Points
        End Function

        ''' <summary>
        ''' Retorna todos os pontos de uma janela retângular dado 2 pontos opostos 2D
        ''' </summary>
        ''' <param name="FirstPoint">Primeiro ponto</param>
        ''' <param name="SecoundPoint">Segundo ponto (oposto)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetWindow2d(FirstPoint As Point2d, SecoundPoint As Point2d) As List(Of Point2d)
            Dim P1 As Point2d
            Dim P2 As Point2d
            Dim P3 As Point2d
            Dim P4 As Point2d
            Dim Points As New List(Of Point2d)
            Points.Add(FirstPoint)
            Points.Add(SecoundPoint)
            P1 = Engine2.Geometry.MinPoint2d(Points)
            P3 = Engine2.Geometry.MaxPoint2d(Points)
            P2 = New Point2d(P3.X, P1.Y)
            P4 = New Point2d(P1.X, P3.Y)
            Points.Clear()
            Points.Add(P1)
            Points.Add(P2)
            Points.Add(P3)
            Points.Add(P4)
            Return Points
        End Function

        ''' <summary>
        ''' Retorna todos os pontos de uma janela retângular dado 2 pontos opostos 3D
        ''' </summary>
        ''' <param name="FirstPoint">Primeiro ponto</param>
        ''' <param name="SecoundPoint">Segundo ponto (oposto)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetWindow3d(FirstPoint As Point3d, SecoundPoint As Point3d) As List(Of Point3d)
            Dim P1 As Point3d
            Dim P2 As Point3d
            Dim P3 As Point3d
            Dim P4 As Point3d
            Dim Points As New List(Of Point3d)
            Points.Add(FirstPoint)
            Points.Add(SecoundPoint)
            P1 = Engine2.Geometry.MinPoint3d(Points)
            P3 = Engine2.Geometry.MaxPoint3d(Points)
            P2 = New Point3d(P3.X, P1.Y, 0)
            P4 = New Point3d(P1.X, P3.Y, 0)
            Points.Clear()
            Points.Add(P1)
            Points.Add(P2)
            Points.Add(P3)
            Points.Add(P4)
            Return Points
        End Function

        ''' <summary>
        ''' Retorna as informações de perpendicular dado uma entidade e um ponto de referência
        ''' </summary>
        '''<param name="Entity">Entidade</param>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <returns>Nothing\PerpendicularInfor</returns>
        ''' <remarks></remarks>
        Public Shared Function GetPerpendicular(Entity As Entity, PointReference As Point3d) As PerpendicularInfor
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim Curve As Curve
            Dim Extension As Double
            Dim Point3d1a As Point3d
            Dim Point3d2a As Point3d
            Dim Point3d1 As Point3d
            Dim Point3d2 As Point3d
            Dim Vector3d1 As Vector3d
            Dim Vector3d2 As Vector3d
            Dim Angle As Double
            Dim OriginalCurve As Curve
            Try
                If Engine2.EntityInteration.IsCurved(Entity) = True Then
                    Curve = Entity
                    Point3d1 = PointReference.TransformBy(Editor.CurrentUserCoordinateSystem)
                    Point3d2 = Curve.GetClosestPointTo(Point3d1, False)
                    Vector3d1 = Point3d1.GetVectorTo(Point3d2)
                    Vector3d2 = Curve.GetFirstDerivative(Point3d2)
                    If Vector3d1.IsPerpendicularTo(Vector3d2) = True Then
                        If (Curve.StartPoint <> Curve.EndPoint) And (Curve.StartPoint <> PointReference) And (Curve.EndPoint <> PointReference) Then
                            If Curve.StartPoint.DistanceTo(PointReference) <= Curve.EndPoint.DistanceTo(PointReference) Then
                                GetPerpendicular = New PerpendicularInfor(False, Point3d2, Curve.StartPoint, PointReference, Curve)
                            Else
                                GetPerpendicular = New PerpendicularInfor(False, Point3d2, Curve.EndPoint, PointReference, Curve)
                            End If
                            Return GetPerpendicular
                        Else
                            Return Nothing
                        End If
                    Else
                        If Curve.StartPoint.DistanceTo(PointReference) <= Curve.EndPoint.DistanceTo(PointReference) Then
                            Point3d1a = Curve.StartPoint
                            Extension = Curve.StartPoint.DistanceTo(PointReference)
                            Angle = Engine2.Geometry.GetAngle(Curve.EndPoint, Curve.StartPoint, eAngleFormat.Degrees)
                        Else
                            Point3d1a = Curve.EndPoint
                            Extension = Curve.EndPoint.DistanceTo(PointReference)
                            Angle = Engine2.Geometry.GetAngle(Curve.StartPoint, Curve.EndPoint, eAngleFormat.Degrees)
                        End If
                        Point3d2a = Engine2.Geometry.PolarPoint3d(Point3d1a, Angle, Extension)
                        OriginalCurve = Entity
                        Curve = New Line(Point3d1a, Point3d2a)
                        Point3d2 = Curve.GetClosestPointTo(Point3d1, False)
                        Vector3d1 = Point3d1.GetVectorTo(Point3d2)
                        Vector3d2 = Curve.GetFirstDerivative(Point3d2)
                        If Vector3d1.IsPerpendicularTo(Vector3d2) = True Then
                            If (Curve.StartPoint <> Curve.EndPoint) And (Curve.StartPoint <> PointReference) And (Curve.EndPoint <> PointReference) Then
                                If OriginalCurve.StartPoint.DistanceTo(PointReference) <= OriginalCurve.EndPoint.DistanceTo(PointReference) Then
                                    GetPerpendicular = New PerpendicularInfor(True, Point3d2, OriginalCurve.StartPoint, PointReference, OriginalCurve)
                                Else
                                    GetPerpendicular = New PerpendicularInfor(True, Point3d2, OriginalCurve.EndPoint, PointReference, OriginalCurve)
                                End If
                                Return GetPerpendicular
                            Else
                                Return Nothing
                            End If
                        Else
                            Return Nothing
                        End If
                    End If
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retorna as informações de perpendicular dado dois pontos referentes ao ponto inicial e final de um segmento e um ponto de referência
        ''' </summary>
        ''' <param name="StartPointInSegment">Primeiro ponto do segmento</param>
        ''' <param name="EndPointInSegment">Ponto final do segmento</param>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <returns>Nothing\PerpendicularInfor</returns>
        ''' <remarks></remarks>
        Public Shared Function GetPerpendicular(StartPointInSegment As Point3d, EndPointInSegment As Point3d, PointReference As Point3d) As PerpendicularInfor
            Dim Database As Database = HostApplicationServices.WorkingDatabase
            Dim Editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim Curve As Curve
            Dim Extension As Double
            Dim Point3d1a As Point3d
            Dim Point3d2a As Point3d
            Dim Point3d1 As Point3d
            Dim Point3d2 As Point3d
            Dim Vector3d1 As Vector3d
            Dim Vector3d2 As Vector3d
            Dim Angle As Double
            Dim OriginalCurve As Curve
            Try
                Curve = New Line(StartPointInSegment, EndPointInSegment)
                Point3d1 = PointReference.TransformBy(Editor.CurrentUserCoordinateSystem)
                Point3d2 = Curve.GetClosestPointTo(Point3d1, False)
                Vector3d1 = Point3d1.GetVectorTo(Point3d2)
                Vector3d2 = Curve.GetFirstDerivative(Point3d2)
                If Vector3d1.IsPerpendicularTo(Vector3d2) = True Then
                    If (Curve.StartPoint <> Curve.EndPoint) And (Curve.StartPoint <> PointReference) And (Curve.EndPoint <> PointReference) Then
                        If Curve.StartPoint.DistanceTo(PointReference) <= Curve.EndPoint.DistanceTo(PointReference) Then
                            GetPerpendicular = New PerpendicularInfor(False, Point3d2, Curve.StartPoint, PointReference, Curve)
                        Else
                            GetPerpendicular = New PerpendicularInfor(False, Point3d2, Curve.EndPoint, PointReference, Curve)
                        End If
                        Return GetPerpendicular
                    Else
                        Return Nothing
                    End If
                Else
                    If Curve.StartPoint.DistanceTo(PointReference) <= Curve.EndPoint.DistanceTo(PointReference) Then
                        Point3d1a = Curve.StartPoint
                        Extension = Curve.StartPoint.DistanceTo(PointReference)
                        Angle = Engine2.Geometry.GetAngle(Curve.EndPoint, Curve.StartPoint, eAngleFormat.Degrees)
                    Else
                        Point3d1a = Curve.EndPoint
                        Extension = Curve.EndPoint.DistanceTo(PointReference)
                        Angle = Engine2.Geometry.GetAngle(Curve.StartPoint, Curve.EndPoint, eAngleFormat.Degrees)
                    End If
                    Point3d2a = Engine2.Geometry.PolarPoint3d(Point3d1a, Angle, Extension)
                    OriginalCurve = New Line(StartPointInSegment, EndPointInSegment)
                    Curve = New Line(Point3d1a, Point3d2a)
                    Point3d2 = Curve.GetClosestPointTo(Point3d1, False)
                    Vector3d1 = Point3d1.GetVectorTo(Point3d2)
                    Vector3d2 = Curve.GetFirstDerivative(Point3d2)
                    If Vector3d1.IsPerpendicularTo(Vector3d2) = True Then
                        If (Curve.StartPoint <> Curve.EndPoint) And (Curve.StartPoint <> PointReference) And (Curve.EndPoint <> PointReference) Then
                            If OriginalCurve.StartPoint.DistanceTo(PointReference) <= OriginalCurve.EndPoint.DistanceTo(PointReference) Then
                                GetPerpendicular = New PerpendicularInfor(True, Point3d2, OriginalCurve.StartPoint, PointReference, OriginalCurve)
                            Else
                                GetPerpendicular = New PerpendicularInfor(True, Point3d2, OriginalCurve.EndPoint, PointReference, OriginalCurve)
                            End If
                            Return GetPerpendicular
                        Else
                            Return Nothing
                        End If
                    Else
                        Return Nothing
                    End If
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha espelhamento de entidades
        ''' </summary>
        ''' <param name="Point3d">Ponto a ser espelhado</param>
        ''' <param name="LineSegment3d">Linha que corresponde a mediatrix</param>
        ''' <returns>Point3dCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function MirrorPoint(Point3d As Point3d, LineSegment3d As LineSegment3d) As Point3d
            Dim OriginalAngle As Double
            Dim MidAngle As Double
            Dim MirrorAngle As Double
            Dim Distance As Double
            OriginalAngle = Engine2.Geometry.GetAngle(LineSegment3d.StartPoint, Point3d, eAngleFormat.Degrees)
            MidAngle = Engine2.Geometry.GetAngle(LineSegment3d.StartPoint, LineSegment3d.EndPoint, eAngleFormat.Degrees)
            MirrorAngle = Engine2.Geometry.MirrorAngle(OriginalAngle, MidAngle)
            Distance = LineSegment3d.StartPoint.DistanceTo(Point3d)
            Return Engine2.Geometry.PolarPoint3d(LineSegment3d.StartPoint, MirrorAngle, Distance)
        End Function

    End Class

    ''' <summary>
    ''' Retorno da função GetPerpendicular
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class PerpendicularInfor

        ''' <summary>
        ''' Retorna o ponto perpendicular
        ''' </summary>
        ''' <remarks></remarks>
        Private _PerpendicularPoint As Point3d

        ''' <summary>
        ''' Retorna o ponto base da projeção
        ''' </summary>
        ''' <remarks></remarks>
        Private _ProjectionBasePoint As Point3d

        ''' <summary>
        ''' Retorna se a perpendicular esta dentro ou fora do segmento
        ''' </summary>
        ''' <remarks></remarks>
        Private _VirtualPerpendicular As Boolean

        ''' <summary>
        ''' Ponto de origem para perpendicular
        ''' </summary>
        ''' <remarks></remarks>
        Private _Origin As Point3d

        ''' <summary>
        ''' Entidade de origem
        ''' </summary>
        ''' <remarks></remarks>
        Private _Curve As Curve

        ''' <summary>
        ''' Variável para uso do usuário
        ''' </summary>
        ''' <remarks></remarks>
        Public Tag As Object

        ''' <summary>
        ''' Retorna o ponto perpendicular
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property PerpendicularPoint As Point3d
            Get
                Return Me._PerpendicularPoint
            End Get
        End Property

        ''' <summary>
        ''' Retorna o ponto base da projeção
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ProjectionBasePoint As Point3d
            Get
                Return Me._ProjectionBasePoint
            End Get
        End Property

        ''' <summary>
        ''' Retorna se a perpendicular esta dentro ou fora do segmento
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property VirtualPerpendicular As Boolean
            Get
                Return Me._VirtualPerpendicular
            End Get
        End Property

        ''' <summary>
        ''' Retorna a curva de origem
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Curve As Curve
            Get
                Return Me._Curve
            End Get
        End Property

        ''' <summary>
        ''' Retorna o ponto base da projeção
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Origin As Point3d
            Get
                Return Me._Origin
            End Get
        End Property

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="VirtualPerpendicular">Informa se a perpendicular é virtual</param>
        ''' <param name="PerpendicularPoint">Ponto perpendicular</param>
        ''' <param name="ProjectionBasePoint">Ponto base da projeção</param>
        ''' <param name="Origin">Ponto de origem para perpendicular</param>
        ''' <param name="Curve">Curva de origem</param>
        ''' <param name="Tag">Variável para uso do usuário</param>
        ''' <remarks></remarks>
        Public Sub New(VirtualPerpendicular As Boolean, PerpendicularPoint As Point3d, ProjectionBasePoint As Point3d, Origin As Point3d, Curve As Curve, Optional Tag As Object = Nothing)
            Me._PerpendicularPoint = PerpendicularPoint
            Me._ProjectionBasePoint = ProjectionBasePoint
            Me._VirtualPerpendicular = VirtualPerpendicular
            Me._Origin = Origin
            Me._Curve = Curve
            Me.Tag = Tag
        End Sub

    End Class


    ''' <summary>
    ''' Função de apoio para ordenar coleção de pontos 2D
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class Sort2dByX : Implements IComparer(Of Point2d)

        ''' <summary>
        ''' Retorna se o argumento é zero
        ''' </summary>
        ''' <param name="a"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsZero(a As Double) As Boolean
            Return Math.Abs(a) < Tolerance.[Global].EqualPoint
        End Function

        ''' <summary>
        ''' Retorna se os argumentos são iguais
        ''' </summary>
        ''' <param name="a"></param>
        ''' <param name="b"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsEqual(a As Double, b As Double) As Boolean
            Return IsZero(b - a)
        End Function

        ''' <summary>
        ''' Compara os pontos
        ''' </summary>
        ''' <param name="a"></param>
        ''' <param name="b"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Compare(a As Point2d, b As Point2d) As Integer Implements IComparer(Of Point2d).Compare
            If IsEqual(a.X, b.X) Then
                Return 0
            End If
            If a.X < b.X Then
                Return -1
            End If
            Return 1
        End Function

    End Class

    ''' <summary>
    ''' Função de apoio para ordenar coleção de pontos 3D
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class Sort3dByX : Implements IComparer(Of Point3d)

        ''' <summary>
        ''' Retorna se o argumento é zero
        ''' </summary>
        ''' <param name="a"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsZero(a As Double) As Boolean
            Return Math.Abs(a) < Tolerance.Global.EqualPoint
        End Function

        ''' <summary>
        ''' Retorna se os argumentos são iguais
        ''' </summary>
        ''' <param name="a"></param>
        ''' <param name="b"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsEqual(a As Double, b As Double) As Boolean
            Return IsZero(b - a)
        End Function

        ''' <summary>
        ''' Compara os pontos
        ''' </summary>
        ''' <param name="a"></param>
        ''' <param name="b"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Compare(a As Point3d, b As Point3d) As Integer Implements IComparer(Of Point3d).Compare
            If IsEqual(a.X, b.X) Then
                Return 0
            End If
            If a.X < b.X Then
                Return -1
            End If
            Return 1
        End Function

    End Class

    ''' <summary>
    ''' Permite cálculos entre pontos polares
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Polar

        ''' <summary>
        ''' Ponto 3D
        ''' </summary>
        ''' <remarks></remarks>
        Private _Point3d As Point3d

        ''' <summary>
        ''' Ângulo em graus
        ''' </summary>
        ''' <remarks></remarks>
        Private _DegreeAngle As Double

        ''' <summary>
        ''' Ângulo em radianos
        ''' </summary>
        ''' <remarks></remarks>
        Private _RadianAngle As Double

        ''' <summary>
        ''' Ponto 3d
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Point3d As Point3d
            Get
                Return Me._Point3d
            End Get
            Set(value As Point3d)
                Me._Point3d = value
            End Set
        End Property

        ''' <summary>
        ''' Ângulo de graus
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DegreeAngle As Double
            Get
                Return Me._DegreeAngle
            End Get
            Set(value As Double)
                Me._DegreeAngle = value
                Me._RadianAngle = Engine2.Geometry.DegreeToRadian(Me._DegreeAngle)
            End Set
        End Property

        ''' <summary>
        ''' Ângulo em radianos
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RadianAngle As Double
            Get
                Return Me._RadianAngle
            End Get
            Set(value As Double)
                Me._RadianAngle = value
                Me._DegreeAngle = Engine2.Geometry.RadianToDegree(Me._RadianAngle)
            End Set
        End Property

        ' ''' <summary>
        ' ''' Obtem os pontos 3d das intersecções com o bloco
        ' ''' </summary>
        ' ''' <param name="BlockReference">BlockReference</param>
        ' ''' <param name="IntersectType">Tipo de busca</param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Function GetIntersectionPoints3dAt(BlockReference As BlockReference, FilterEntity As ArrayList, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.OnBothOperands) As Point3dCollection
        '    Try

        '        'Obtem os pontos de intersecção
        '        Return Engine2.Geometry.VirtualIntersection(Me._Point3d, Me._DegreeAngle, BlockReference, FilterEntity, , IntersectType, , , False)

        '    Catch

        '        'Erro
        '        Return Nothing

        '    End Try
        'End Function

        ' ''' <summary>
        ' ''' Obtem os pontos 2d das intersecções com o bloco
        ' ''' </summary>
        ' ''' <param name="BlockReference">BlockReference</param>
        ' ''' <param name="FilterEntity">Filtro de entidade</param>
        ' ''' <param name="IntersectType">Tipo de busca</param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        'Public Function GetIntersectionPoints2dAt(BlockReference As BlockReference, FilterEntity As ArrayList, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.OnBothOperands) As Point2dCollection
        '    Try

        '        'Prepara a variável de retorno
        '        GetIntersectionPoints2dAt = New Point2dCollection

        '        'Declarações
        '        Dim Point3dCollection As Point3dCollection

        '        'Obtem os pontos de intersecção
        '        Point3dCollection = Engine2.Geometry.VirtualIntersection(Me._Point3d, Me._DegreeAngle, BlockReference, FilterEntity, , IntersectType, , , False)

        '        'Avalia o retorno
        '        If Point3dCollection.Count > 0 Then

        '            'Percorre a coleção
        '            For Each Point3d As Point3d In Point3dCollection

        '                'Monta a coleção de pontos 2d
        '                GetIntersectionPoints2dAt.Add(Engine2.Geometry.Point3dToPoint2d(Point3d))

        '            Next

        '            'Retorno
        '            Return GetIntersectionPoints2dAt

        '        Else

        '            'Retorno
        '            Return GetIntersectionPoints2dAt

        '        End If

        '    Catch

        '        'Erro
        '        Return Nothing

        '    End Try
        'End Function

        ''' <summary>
        ''' Obtem os pontos 3d das intersecções com a curva
        ''' </summary>
        ''' <param name="Curve">Curve</param>
        ''' <param name="IntersectType">Tipo de busca</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetIntersectionPoints3dAt(Curve As Curve, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.OnBothOperands) As Point3dCollection
            Try

                'Declarações
                Dim IntersectionEntityCollection As IntersectionEntityCollection
                Dim TypedValueCollection As New List(Of TypedValue)
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim Line As Line
                Dim BlockTableRecord As BlockTableRecord
                Dim Distance As Double

                'Abre a transação
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction

                        'Abre BlockTableRecord
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)

                        'Obtem a distancia do segmento
                        Distance = Me._Point3d.DistanceTo(Curve.GeometricExtents.MaxPoint)

                        'Cria a linha
                        Line = New Line(Me.Point3d, Engine2.Geometry.PolarPoint3d(Me._Point3d, Me._DegreeAngle, (Distance * 2)))

                        'Anexa a linha em BlockTableRecord
                        BlockTableRecord.AppendEntity(Line)

                        'Cria a linha 
                        Transaction.AddNewlyCreatedDBObject(Line, True)

                        'Monta o filtro
                        With TypedValueCollection
                            .Clear()
                            .Add(New TypedValue(DxfCode.Handle, Line.Handle))
                        End With

                        'Calcula as intersecções
                        IntersectionEntityCollection = Engine2.EntityInteration.GetEntityIntersections(Curve, IntersectType, TypedValueCollection)

                        'Apaga a linha
                        Line.Erase()

                        'Confirma a transação
                        Transaction.Commit()

                    End Using

                End Using

                'Avalia o retorno
                If IsNothing(IntersectionEntityCollection) = True Then

                    'Retorno
                    Return New Point3dCollection

                Else

                    'Retorno
                    Return New Point3dCollection(IntersectionEntityCollection.IntersectionPoints3d.ToArray)

                End If

            Catch

                'Erro
                Return Nothing

            End Try
        End Function

        ''' <summary>
        ''' Obtem os pontos 2d das intersecções com a curva
        ''' </summary>
        ''' <param name="Curve">Curve</param>
        ''' <param name="IntersectType">Tipo de busca</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetIntersectionPoints2dAt(Curve As Curve, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.OnBothOperands) As Point2dCollection
            Try

                'Declarações
                Dim IntersectionEntityCollection As IntersectionEntityCollection
                Dim TypedValueCollection As New List(Of TypedValue)
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim Line As Line
                Dim BlockTableRecord As BlockTableRecord
                Dim Distance As Double

                'Abre a transação
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction

                        'Abre BlockTableRecord
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)

                        'Obtem a distancia do segmento
                        Distance = Me._Point3d.DistanceTo(Curve.GeometricExtents.MaxPoint)

                        'Cria a linha
                        Line = New Line(Me.Point3d, Engine2.Geometry.PolarPoint3d(Me._Point3d, Me._DegreeAngle, (Distance * 2)))

                        'Anexa a linha em BlockTableRecord
                        BlockTableRecord.AppendEntity(Line)

                        'Cria a linha 
                        Transaction.AddNewlyCreatedDBObject(Line, True)

                        'Monta o filtro
                        With TypedValueCollection
                            .Clear()
                            .Add(New TypedValue(DxfCode.Handle, Line.Handle))
                        End With

                        'Calcula as intersecções
                        IntersectionEntityCollection = Engine2.EntityInteration.GetEntityIntersections(Curve, IntersectType, TypedValueCollection)

                        'Apaga a linha
                        Line.Erase()

                        'Confirma a transação
                        Transaction.Commit()

                    End Using

                End Using

                'Avalia o retorno
                If IsNothing(IntersectionEntityCollection) = True Then

                    'Retorno
                    Return New Point2dCollection

                Else

                    'Retorno
                    Return New Point2dCollection(IntersectionEntityCollection.IntersectionPoints2d.ToArray)

                End If

            Catch

                'Erro
                Return Nothing

            End Try
        End Function

        ''' <summary>
        ''' Obtem as entidades e pontos de intersecção 
        ''' </summary>
        ''' <param name="Line2d">Line2d</param>
        ''' <param name="IntersectType">Tipo de busca</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetIntersectionPoint3dAt(Line2d As Line2d, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendBoth) As Point3d
            Dim Point3dCollection As Point3dCollection
            Point3dCollection = Engine2.Geometry.VirtualIntersection(Me.Point3d, Me._DegreeAngle, Line2d, IntersectType)
            If Point3dCollection.Count > 0 Then
                Return Point3dCollection.Item(0)
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Obtem o ponto 3d da intersecção
        ''' </summary>
        ''' <param name="LineSegment2d">LineSegment2d</param>
        ''' <param name="IntersectType">Tipo de busca</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetIntersectionPoint3dAt(Linesegment2d As LineSegment2d, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendBoth) As Point3d
            Dim Point3dCollection As Point3dCollection
            Point3dCollection = Engine2.Geometry.VirtualIntersection(Me.Point3d, Me._DegreeAngle, Linesegment2d, IntersectType)
            If Point3dCollection.Count > 0 Then
                Return Point3dCollection.Item(0)
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Obtem o ponto 3d da intersecção
        ''' </summary>
        ''' <param name="Polar">Polar</param>
        ''' <param name="IntersectType">Tipo de busca</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetIntersectionPoint3dAt(Polar As Polar, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendBoth) As Point3d
            Dim Point3dCollection As Point3dCollection
            Point3dCollection = Engine2.Geometry.VirtualIntersection(New LineSegment3d(Me._Point3d, Engine2.Geometry.PolarPoint3d(Me._Point3d, Me._DegreeAngle, 1)), New LineSegment3d(Polar.Point3d, Engine2.Geometry.PolarPoint3d(Polar.Point3d, Polar.DegreeAngle, 1)), , IntersectType)
            If Point3dCollection.Count > 0 Then
                Return Point3dCollection.Item(0)
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Obtem o ponto 2d da intersecção
        ''' </summary>
        ''' <param name="Polar">Polar</param>
        ''' <param name="IntersectType">Tipo debusca</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetIntersectionPoint2dAt(Polar As Polar, Optional IntersectType As Autodesk.AutoCAD.DatabaseServices.Intersect = Intersect.ExtendBoth) As Point2d
            Dim Point3dCollection As Point3dCollection
            Point3dCollection = Engine2.Geometry.VirtualIntersection(New LineSegment3d(Me._Point3d, Engine2.Geometry.PolarPoint3d(Me._Point3d, Me._DegreeAngle, 1)), New LineSegment3d(Polar.Point3d, Engine2.Geometry.PolarPoint3d(Polar.Point3d, Polar.DegreeAngle, 1)), , IntersectType)
            If Point3dCollection.Count > 0 Then
                Return Engine2.Geometry.Point3dToPoint2d(Point3dCollection.Item(0))
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Obtem o ponto 3d na distância
        ''' </summary>
        ''' <param name="Distance">Distância</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPoint3dAtDistance(Distance As Double) As Point3d
            Return Engine2.Geometry.PolarPoint3d(Me._Point3d, Me._DegreeAngle, Distance)
        End Function

        ''' <summary>
        ''' Obtem o ponto 2d na intersecção
        ''' </summary>
        ''' <param name="Distance">Distância</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPoint2dAtDistance(Distance As Double) As Point2d
            Return Engine2.Geometry.Point3dToPoint2d(Engine2.Geometry.PolarPoint3d(Me._Point3d, Me._DegreeAngle, Distance))
        End Function

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="Point3d"></param>
        ''' <param name="Angle"></param>
        ''' <param name="AngleFormat"></param>
        ''' <remarks></remarks>
        Public Sub New(Point3d As Point3d, Angle As Double, Optional AngleFormat As eAngleFormat = eAngleFormat.Degrees)
            Me.Point3d = Point3d
            Select Case AngleFormat
                Case eAngleFormat.Degrees
                    Me.DegreeAngle = Angle
                Case eAngleFormat.Radian
                    Me.RadianAngle = Angle
            End Select
        End Sub

    End Class

End Namespace