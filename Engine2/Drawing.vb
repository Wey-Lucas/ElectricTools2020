Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports System.Text
Imports Autodesk.AutoCAD
Imports System.Runtime.InteropServices
Imports System.Drawing

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Classe de desenho
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Drawing

        ''' <summary>
        ''' Opções de caracteres para desenho de vetor
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eVectorChar
            ''' <summary>
            ''' Zero
            ''' </summary>
            ''' <remarks></remarks>
            _0 = 0

            ''' <summary>
            ''' Um
            ''' </summary>
            ''' <remarks></remarks>
            _1 = 1

            ''' <summary>
            ''' Dois
            ''' </summary>
            ''' <remarks></remarks>
            _2 = 2

            ''' <summary>
            ''' Três
            ''' </summary>
            ''' <remarks></remarks>
            _3 = 3

            ''' <summary>
            ''' Quatro
            ''' </summary>
            ''' <remarks></remarks>
            _4 = 4

            ''' <summary>
            ''' Cinco
            ''' </summary>
            ''' <remarks></remarks>
            _5 = 5

            ''' <summary>
            ''' Seis
            ''' </summary>
            ''' <remarks></remarks>
            _6 = 6

            ''' <summary>
            ''' Sete
            ''' </summary>
            ''' <remarks></remarks>
            _7 = 7

            ''' <summary>
            ''' Oito
            ''' </summary>
            ''' <remarks></remarks>
            _8 = 8

            ''' <summary>
            ''' Nove
            ''' </summary>
            ''' <remarks></remarks>
            _9 = 9

            ''' <summary>
            ''' Ponto
            ''' </summary>
            ''' <remarks></remarks>
            PONTO = 10

            ''' <summary>
            ''' Vírgula
            ''' </summary>
            ''' <remarks></remarks>
            VIRGULA = 11

            ''' <summary>
            ''' ABRE PARENTESE
            ''' </summary>
            ''' <remarks></remarks>
            ABRE_PARENTESE = 12

            ''' <summary>
            ''' FECHA PARENTESE
            ''' </summary>
            ''' <remarks></remarks>
            FECHA_PARENTESE = 13

            ''' <summary>
            ''' ABRE COLCHETE
            ''' </summary>
            ''' <remarks></remarks>
            ABRE_COLCHETE = 14

            ''' <summary>
            ''' FECHA COLCHETE
            ''' </summary>
            ''' <remarks></remarks>
            FECHA_COLCHETE = 15

            ''' <summary>
            ''' ABRE CHAVE
            ''' </summary>
            ''' <remarks></remarks>
            ABRE_CHAVE = 16

            ''' <summary>
            ''' FECHA CHAVE
            ''' </summary>
            ''' <remarks></remarks>
            FECHA_CHAVE = 17

            ''' <summary>
            ''' SOMA
            ''' </summary>
            ''' <remarks></remarks>
            SOMA = 18

            ''' <summary>
            ''' DIVISÃO
            ''' </summary>
            ''' <remarks></remarks>
            DIVISAO = 19

            ''' <summary>
            ''' MULTIPLICAÇÃO
            ''' </summary>
            ''' <remarks></remarks>
            MULTIPLICACAO = 20

            ''' <summary>
            ''' SUBTRAÇÃO
            ''' </summary>
            ''' <remarks></remarks>
            SUBTRACAO = 21

            ''' <summary>
            ''' MAIOR
            ''' </summary>
            ''' <remarks></remarks>
            MAIOR = 22

            ''' <summary>
            ''' MENOR
            ''' </summary>
            ''' <remarks></remarks>
            MENOR = 23

            ''' <summary>
            ''' IGUAL
            ''' </summary>
            ''' <remarks></remarks>
            IGUAL = 24

            ''' <summary>
            ''' DIFERENTE
            ''' </summary>
            ''' <remarks></remarks>
            DIFERENTE = 25

            ''' <summary>
            ''' PORCENTAGEM
            ''' </summary>
            ''' <remarks></remarks>
            PORCENTAGEM = 26

            ''' <summary>
            ''' Letra A
            ''' </summary>
            ''' <remarks></remarks>
            A = 27


            ''' <summary>
            ''' Letra B
            ''' </summary>
            ''' <remarks></remarks>
            B = 28

            ''' <summary>
            ''' Letra C
            ''' </summary>
            ''' <remarks></remarks>
            C = 29

            ''' <summary>
            ''' Letra D
            ''' </summary>
            ''' <remarks></remarks>
            D = 30

            ''' <summary>
            ''' Letra E
            ''' </summary>
            ''' <remarks></remarks>
            E = 31

            ''' <summary>
            ''' Letra F
            ''' </summary>
            ''' <remarks></remarks>
            F = 32

            ''' <summary>
            ''' Letra G
            ''' </summary>
            ''' <remarks></remarks>
            G = 33

            ''' <summary>
            ''' Letra H
            ''' </summary>
            ''' <remarks></remarks>
            H = 34

            ''' <summary>
            ''' Letra I
            ''' </summary>
            ''' <remarks></remarks>
            I = 35

            ''' <summary>
            ''' Letra J
            ''' </summary>
            ''' <remarks></remarks>
            J = 36

            ''' <summary>
            ''' Letra K
            ''' </summary>
            ''' <remarks></remarks>
            K = 37

            ''' <summary>
            ''' Letra L
            ''' </summary>
            ''' <remarks></remarks>
            L = 38

            ''' <summary>
            ''' Letra M
            ''' </summary>
            ''' <remarks></remarks>
            M = 39

            ''' <summary>
            ''' Letra N
            ''' </summary>
            ''' <remarks></remarks>
            N = 40

            ''' <summary>
            ''' Letra O
            ''' </summary>
            ''' <remarks></remarks>
            O = 41

            ''' <summary>
            ''' Letra P
            ''' </summary>
            ''' <remarks></remarks>
            P = 42

            ''' <summary>
            ''' Letra Q
            ''' </summary>
            ''' <remarks></remarks>
            Q = 43

            ''' <summary>
            ''' Letra R
            ''' </summary>
            ''' <remarks></remarks>
            R = 44

            ''' <summary>
            ''' Letra S
            ''' </summary>
            ''' <remarks></remarks>
            S = 45

            ''' <summary>
            ''' Letra T
            ''' </summary>
            ''' <remarks></remarks>
            T = 46

            ''' <summary>
            ''' Letra U
            ''' </summary>
            ''' <remarks></remarks>
            U = 47

            ''' <summary>
            ''' Letra V
            ''' </summary>
            ''' <remarks></remarks>
            V = 48

            ''' <summary>
            ''' Letra X
            ''' </summary>
            ''' <remarks></remarks>
            X = 49

            ''' <summary>
            ''' Letra Y
            ''' </summary>
            ''' <remarks></remarks>
            Y = 50

            ''' <summary>
            ''' Letra W
            ''' </summary>
            ''' <remarks></remarks>
            W = 51

            ''' <summary>
            ''' Letra Z
            ''' </summary>
            ''' <remarks></remarks>
            Z = 52

        End Enum

        ''' <summary>
        ''' Retorna o valor de tolerância mínimo para a criação de entidades gráficas
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetDrawTolerance() As Double
            Dim Units As Integer = Application.GetSystemVariable("LUPREC")
            Dim Value As String
            If Units = 0 Then
                Return 1
            Else
                Value = "0."
                For Index As Integer = 1 To Units
                    If Index = Units Then
                        Value = Value & "1"
                    Else
                        Value = Value & "0"
                    End If
                Next
                Return CDbl(Value)
            End If
        End Function

        ''' <summary>
        ''' Desenha um retângulo
        ''' </summary>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="EndPoint">Ponto final</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Polyline</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawRectangle(StartPoint As Point3d, EndPoint As Point3d, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Polyline
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim Points As List(Of Point3d) = Engine2.Geometry.GetWindow3d(StartPoint, EndPoint)
                Dim BlockTableRecord As BlockTableRecord
                Dim Polyline As New Polyline
                With Polyline
                    For Index As Integer = 0 To Points.Count - 1
                        .AddVertexAt(Index, Engine2.Geometry.Point3dToPoint2d(Points.Item(Index)), 0, 0, 0)
                    Next
                    .Closed = True
                    If IsNothing(Layer) = False Then
                        .Layer = Layer
                    End If
                    If UseUCS = True Then
                        .TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                End With
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        Try
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
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha um retângulo
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="EndPoint">Ponto final</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar a UCS</param>
        ''' <returns>Polyline</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawRectangle(Transaction As Transaction, StartPoint As Point3d, EndPoint As Point3d, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Polyline
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim Points As List(Of Point3d) = Engine2.Geometry.GetWindow3d(StartPoint, EndPoint)
                Dim BlockTableRecord As BlockTableRecord
                Dim Polyline As New Polyline
                With Polyline
                    For Index As Integer = 0 To Points.Count - 1
                        .AddVertexAt(Index, Engine2.Geometry.Point3dToPoint2d(Points.Item(Index)), 0, 0, 0)
                    Next
                    .Closed = True
                    If IsNothing(Layer) = False Then
                        .Layer = Layer
                    End If
                    If UseUCS = True Then
                        .TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                End With
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                BlockTableRecord.AppendEntity(Polyline)
                Transaction.AddNewlyCreatedDBObject(Polyline, True)
                Return Polyline
            Catch
                Return Nothing
            End Try
        End Function


        ''' <summary>
        ''' Desenhar linha
        ''' </summary>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="EndPoint">Ponto final</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Line</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawLine(StartPoint As Point3d, EndPoint As Point3d, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Line
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim Line As Line = Nothing
                Dim retVal As Boolean = False
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        Try
                            BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                            Line = New DatabaseServices.Line(StartPoint, EndPoint)
                            If IsNothing(Layer) = False Then
                                Line.Layer = Layer
                                If UseUCS = True Then
                                    Line.TransformBy(Editor.CurrentUserCoordinateSystem)
                                End If
                            End If
                            BlockTableRecord.AppendEntity(Line)
                            Transaction.AddNewlyCreatedDBObject(Line, True)
                            Transaction.Commit()
                            Return Line
                        Catch
                            Transaction.Abort()
                            Return Nothing
                        End Try
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenhar linha
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="EndPoint">Ponto final</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Line</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawLine(Transaction As Transaction, StartPoint As Point3d, EndPoint As Point3d, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Line
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim Line As Line = Nothing
                Dim retVal As Boolean = False
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                Line = New DatabaseServices.Line(StartPoint, EndPoint)
                If IsNothing(Layer) = False Then
                    Line.Layer = Layer
                    If UseUCS = True Then
                        Line.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                End If
                BlockTableRecord.AppendEntity(Line)
                Transaction.AddNewlyCreatedDBObject(Line, True)
                Return Line
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenhar Arco
        ''' </summary>
        ''' <param name="Center">Centro</param>
        ''' <param name="Radius">Raio</param>
        ''' <param name="StartAngle">Ângulo inicial</param>
        ''' <param name="EndAngle">Ângulo final</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Arc</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawArc(Center As Point3d, Radius As Double, StartAngle As Double, EndAngle As Double, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Arc
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim BlockTable As BlockTable
                Dim Arc As Arc = Nothing
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        Try
                            BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                            BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                            Arc = New Arc(Center, Radius, StartAngle, EndAngle)
                            If IsNothing(Layer) = False Then
                                Arc.Layer = Layer
                                If UseUCS = True Then
                                    Arc.TransformBy(Editor.CurrentUserCoordinateSystem)
                                End If
                            End If
                            Arc.SetDatabaseDefaults()
                            BlockTableRecord.AppendEntity(Arc)
                            Transaction.AddNewlyCreatedDBObject(Arc, True)
                            Transaction.Commit()
                            Return Arc
                        Catch ex As System.Exception
                            Transaction.Abort()
                            Return Nothing
                        End Try
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenhar Arco
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Center">Centro</param>
        ''' <param name="Radius">Raio</param>
        ''' <param name="StartAngle">Ângulo inicial</param>
        ''' <param name="EndAngle">Ângulo final</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Arc</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawArc(Transaction As Transaction, Center As Point3d, Radius As Double, StartAngle As Double, EndAngle As Double, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Arc
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim BlockTable As BlockTable
                Dim Arc As Arc = Nothing
                BlockTable = Transaction.GetObject(Database.BlockTableId, OpenMode.ForRead)
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                Arc = New Arc(Center, Radius, StartAngle, EndAngle)
                If IsNothing(Layer) = False Then
                    Arc.Layer = Layer
                    If UseUCS = True Then
                        Arc.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                End If
                Arc.SetDatabaseDefaults()
                BlockTableRecord.AppendEntity(Arc)
                Transaction.AddNewlyCreatedDBObject(Arc, True)
                Return Arc
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha arco com 3 pontos
        ''' </summary>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="PointAtArc">Ponto no arco</param>
        ''' <param name="EndPoint">Ponto final</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Arc</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawArc(StartPoint As Point3d, PointAtArc As Point3d, EndPoint As Point3d, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Arc
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim CircularArc3d As New CircularArc3d(StartPoint, PointAtArc, EndPoint)
                Dim Arc As Arc
                Dim Center As Point3d = CircularArc3d.Center
                Dim Normal As Vector3d = CircularArc3d.Normal
                Dim Reference As Vector3d = CircularArc3d.ReferenceVector
                Dim Plane As New Plane(Center, Normal)
                Dim Angle As Double = Reference.AngleOnPlane(Plane)
                Arc = New Arc(Center, Normal, CircularArc3d.Radius, CircularArc3d.StartAngle + Angle, CircularArc3d.EndAngle + Angle)
                If IsNothing(Layer) = False Then
                    Arc.Layer = Layer
                    If UseUCS = True Then
                        Arc.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                End If
                Arc.SetDatabaseDefaults()
                CircularArc3d.Dispose()
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        Try
                            BlockTableRecord = DirectCast(Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                            BlockTableRecord.AppendEntity(Arc)
                            Transaction.AddNewlyCreatedDBObject(Arc, True)
                            Transaction.Commit()
                            Return Arc
                        Catch
                            Transaction.Abort()
                            Return Nothing
                        End Try
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha arco com 3 pontos
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="PointAtArc">Ponto no arco</param>
        ''' <param name="EndPoint">Ponto final</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Arc</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawArc(Transaction As Transaction, StartPoint As Point3d, PointAtArc As Point3d, EndPoint As Point3d, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Arc
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim CircularArc3d As New CircularArc3d(StartPoint, PointAtArc, EndPoint)
                Dim Arc As Arc
                Dim Center As Point3d = CircularArc3d.Center
                Dim Normal As Vector3d = CircularArc3d.Normal
                Dim Reference As Vector3d = CircularArc3d.ReferenceVector
                Dim Plane As New Plane(Center, Normal)
                Dim Angle As Double = Reference.AngleOnPlane(Plane)
                Arc = New Arc(Center, Normal, CircularArc3d.Radius, CircularArc3d.StartAngle + Angle, CircularArc3d.EndAngle + Angle)
                If IsNothing(Layer) = False Then
                    Arc.Layer = Layer
                    If UseUCS = True Then
                        Arc.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                End If
                Arc.SetDatabaseDefaults()
                CircularArc3d.Dispose()
                BlockTableRecord = DirectCast(Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                BlockTableRecord.AppendEntity(Arc)
                Transaction.AddNewlyCreatedDBObject(Arc, True)
                Return Arc
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenhar Circulo
        ''' </summary>
        ''' <param name="Center">Centro</param>
        ''' <param name="Radius">Raio</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Circle</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawCircle(Center As Point3d, Radius As Double, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Circle
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim Circle As Circle = Nothing
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        Circle = New Circle(Center, Vector3d.ZAxis, Radius)
                        If IsNothing(Layer) = False Then
                            Circle.Layer = Layer
                            If UseUCS = True Then
                                Circle.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                        End If
                        Circle.SetDatabaseDefaults()
                        BlockTableRecord.AppendEntity(Circle)
                        Transaction.AddNewlyCreatedDBObject(Circle, True)
                        Transaction.Commit()
                    End Using
                End Using
                Return Circle
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenhar Circulo
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Center">Centro</param>
        ''' <param name="Radius">Raio</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Circle</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawCircle(Transaction As Transaction, Center As Point3d, Radius As Double, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Circle
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim Circle As Circle = Nothing
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                Circle = New Circle(Center, Vector3d.ZAxis, Radius)
                If IsNothing(Layer) = False Then
                    Circle.Layer = Layer
                    If UseUCS = True Then
                        Circle.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                End If
                Circle.SetDatabaseDefaults()
                BlockTableRecord.AppendEntity(Circle)
                Transaction.AddNewlyCreatedDBObject(Circle, True)
                Return Circle
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha a dimensão alinhada
        ''' </summary>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="Line1Point">Ponto de ancoragem inicial</param>
        ''' <param name="Line2Point">Ponto de ancoragem final</param>
        ''' <param name="OffSet">Distância os pontos de ancoragem</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>RotatedDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawRotatedDimension(DimStyleName As String, ByVal Line1Point As Point3d, ByVal Line2Point As Point3d, ByVal OffSet As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As RotatedDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim Angle As Double
                Dim DimensionLinePoint As Point3d
                Dim RotatedDimension As RotatedDimension
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                        If DimStyleTable.Has(DimStyleName) = True Then
                            DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                            Angle = Engine2.Geometry.GetAngle(Line1Point, Line2Point, Geometry.eAngleFormat.Degrees)
                            DimensionLinePoint = Engine2.Geometry.PolarPoint3d(Line1Point, If(OffSet > 0, Angle + 90, Angle + 270), OffSet)
                            RotatedDimension = New RotatedDimension(Engine2.Geometry.DegreeToRadian(Angle), Line1Point, Line2Point, DimensionLinePoint, DimensionText, DimStyleTableRecord.ObjectId)
                            RotatedDimension.SetDatabaseDefaults()
                            If IsNothing(Layer) = False Then
                                RotatedDimension.Layer = Layer
                            End If
                            If UseUCS = True Then
                                RotatedDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                            BlockTableRecord.AppendEntity(RotatedDimension)
                            Transaction.AddNewlyCreatedDBObject(RotatedDimension, True)
                            Transaction.Commit()
                            Return RotatedDimension
                        Else
                            Transaction.Abort()
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha a dimensão alinhada
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="Line1Point">Ponto de ancoragem inicial</param>
        ''' <param name="Line2Point">Ponto de ancoragem final</param>
        ''' <param name="OffSet">Distância os pontos de ancoragem</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>RotatedDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawRotatedDimension(Transaction As Transaction, DimStyleName As String, ByVal Line1Point As Point3d, ByVal Line2Point As Point3d, ByVal OffSet As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As RotatedDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim Angle As Double
                Dim DimensionLinePoint As Point3d
                Dim RotatedDimension As RotatedDimension
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                If DimStyleTable.Has(DimStyleName) = True Then
                    DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                    Angle = Engine2.Geometry.GetAngle(Line1Point, Line2Point, Geometry.eAngleFormat.Degrees)
                    DimensionLinePoint = Engine2.Geometry.PolarPoint3d(Line1Point, If(OffSet > 0, Angle + 90, Angle + 270), OffSet)
                    RotatedDimension = New RotatedDimension(Engine2.Geometry.DegreeToRadian(Angle), Line1Point, Line2Point, DimensionLinePoint, DimensionText, DimStyleTableRecord.ObjectId)
                    RotatedDimension.SetDatabaseDefaults()
                    If IsNothing(Layer) = False Then
                        RotatedDimension.Layer = Layer
                    End If
                    If UseUCS = True Then
                        RotatedDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    BlockTableRecord.AppendEntity(RotatedDimension)
                    Transaction.AddNewlyCreatedDBObject(RotatedDimension, True)
                    Return RotatedDimension
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha a dimensão alinhada
        ''' </summary>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="Line1Point">Ponto de ancoragem inicial</param>
        ''' <param name="Line2Point">Ponto de cálculo (alinhamento da linha de cota)</param>
        ''' <param name="Line3Point">Ponto de ancoragem final</param>
        ''' <param name="OffSet">Distância os pontos de ancoragem</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>RotatedDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawRotatedDimension(DimStyleName As String, Line1Point As Point3d, Line2Point As Point3d, Line3Point As Point3d, OffSet As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As RotatedDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim Angle As Double
                Dim DimensionLinePoint As Point3d
                Dim RotatedDimension As RotatedDimension
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                        If DimStyleTable.Has(DimStyleName) = True Then
                            DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                            Angle = Engine2.Geometry.GetAngle(Line1Point, Line2Point, Geometry.eAngleFormat.Degrees)
                            DimensionLinePoint = Engine2.Geometry.PolarPoint3d(Line1Point, If(OffSet > 0, Angle + 90, Angle + 270), OffSet)
                            RotatedDimension = New RotatedDimension(Engine2.Geometry.DegreeToRadian(Angle), Line1Point, Line3Point, DimensionLinePoint, DimensionText, DimStyleTableRecord.ObjectId)
                            RotatedDimension.SetDatabaseDefaults()
                            If IsNothing(Layer) = False Then
                                RotatedDimension.Layer = Layer
                            End If
                            If UseUCS = True Then
                                RotatedDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                            BlockTableRecord.AppendEntity(RotatedDimension)
                            Transaction.AddNewlyCreatedDBObject(RotatedDimension, True)
                            Transaction.Commit()
                            Return RotatedDimension
                        Else
                            Transaction.Abort()
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha a dimensão alinhada
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="Line1Point">Ponto de ancoragem inicial</param>
        ''' <param name="Line2Point">Ponto de cálculo (alinhamento da linha de cota)</param>
        ''' <param name="Line3Point">Ponto de ancoragem final</param>
        ''' <param name="OffSet">Distância os pontos de ancoragem</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param> 
        ''' <returns>RotatedDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawRotatedDimension(Transaction As Transaction, DimStyleName As String, Line1Point As Point3d, Line2Point As Point3d, Line3Point As Point3d, OffSet As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As RotatedDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim Angle As Double
                Dim DimensionLinePoint As Point3d
                Dim RotatedDimension As RotatedDimension
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                If DimStyleTable.Has(DimStyleName) = True Then
                    DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                    Angle = Engine2.Geometry.GetAngle(Line1Point, Line2Point, Geometry.eAngleFormat.Degrees)
                    DimensionLinePoint = Engine2.Geometry.PolarPoint3d(Line1Point, If(OffSet > 0, Angle + 90, Angle + 270), OffSet)
                    RotatedDimension = New RotatedDimension(Engine2.Geometry.DegreeToRadian(Angle), Line1Point, Line3Point, DimensionLinePoint, DimensionText, DimStyleTableRecord.ObjectId)
                    RotatedDimension.SetDatabaseDefaults()
                    If IsNothing(Layer) = False Then
                        RotatedDimension.Layer = Layer
                    End If
                    If UseUCS = True Then
                        RotatedDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    BlockTableRecord.AppendEntity(RotatedDimension)
                    Transaction.AddNewlyCreatedDBObject(RotatedDimension, True)
                    Return RotatedDimension
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha a dimensão linear
        ''' </summary>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="Line1Point">Ponto de ancoragem inicial</param>
        ''' <param name="Line2Point">Ponto de ancoragem final</param>
        ''' <param name="OffSet">Distância os pontos de ancoragem</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>RotatedDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawAlignedDimension(DimStyleName As String, Line1Point As Point3d, Line2Point As Point3d, OffSet As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As AlignedDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim Angle As Double
                Dim DimensionLinePoint As Point3d
                Dim AlignedDimension As AlignedDimension
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                        If DimStyleTable.Has(DimStyleName) = True Then
                            DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                            Dim Point3dCollection As New Point3dCollection
                            Point3dCollection.Add(Line1Point)
                            Point3dCollection.Add(Line2Point)
                            Angle = Engine2.Geometry.GetAngle(Engine2.Geometry.MinPoint3d(Point3dCollection), Engine2.Geometry.MaxPoint3d(Point3dCollection), Geometry.eAngleFormat.Degrees)
                            DimensionLinePoint = Engine2.Geometry.PolarPoint3d(Line1Point, If(OffSet > 0, Angle + 90, Angle + 270), OffSet)
                            AlignedDimension = New AlignedDimension(Line1Point, Line2Point, DimensionLinePoint, DimensionText, DimStyleTableRecord.ObjectId)
                            AlignedDimension.SetDatabaseDefaults()
                            If IsNothing(Layer) = False Then
                                AlignedDimension.Layer = Layer
                            End If
                            If UseUCS = True Then
                                AlignedDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                            BlockTableRecord.AppendEntity(AlignedDimension)
                            Transaction.AddNewlyCreatedDBObject(AlignedDimension, True)
                            Transaction.Commit()
                            Return AlignedDimension
                        Else
                            Transaction.Abort()
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha a dimensão linear
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="Line1Point">Ponto de ancoragem inicial</param>
        ''' <param name="Line2Point">Ponto de ancoragem final</param>
        ''' <param name="OffSet">Distância os pontos de ancoragem</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>RotatedDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawAlignedDimension(Transaction As Transaction, DimStyleName As String, Line1Point As Point3d, Line2Point As Point3d, OffSet As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As AlignedDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim Angle As Double
                Dim DimensionLinePoint As Point3d
                Dim AlignedDimension As AlignedDimension
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                If DimStyleTable.Has(DimStyleName) = True Then
                    DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                    Dim Point3dCollection As New Point3dCollection
                    Point3dCollection.Add(Line1Point)
                    Point3dCollection.Add(Line2Point)
                    Angle = Engine2.Geometry.GetAngle(Engine2.Geometry.MinPoint3d(Point3dCollection), Engine2.Geometry.MaxPoint3d(Point3dCollection), Geometry.eAngleFormat.Degrees)
                    DimensionLinePoint = Engine2.Geometry.PolarPoint3d(Line1Point, If(OffSet > 0, Angle + 90, Angle + 270), OffSet)
                    AlignedDimension = New AlignedDimension(Line1Point, Line2Point, DimensionLinePoint, DimensionText, DimStyleTableRecord.ObjectId)
                    AlignedDimension.SetDatabaseDefaults()
                    If IsNothing(Layer) = False Then
                        AlignedDimension.Layer = Layer
                    End If
                    If UseUCS = True Then
                        AlignedDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    BlockTableRecord.AppendEntity(AlignedDimension)
                    Transaction.AddNewlyCreatedDBObject(AlignedDimension, True)
                    Return AlignedDimension
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão angular
        ''' </summary>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="Line1Start">Ponto inicial da primeira linha</param>
        ''' <param name="Line1End">Ponto final da primeira linha</param>
        ''' <param name="Line2Start">Ponto inicial da outra linha</param>
        ''' <param name="Line2End">Ponto final da outra linha</param>
        ''' <param name="ArcPoint">Ponto de fixação</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>LineAngularDimension2</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawLineAngularDimension(DimStyleName As String, Line1Start As Point3d, Line1End As Point3d, Line2Start As Point3d, Line2End As Point3d, ByVal ArcPoint As Point3d, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As LineAngularDimension2
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim LineAngularDimension2 As LineAngularDimension2
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                        If DimStyleTable.Has(DimStyleName) = True Then
                            DimStyleTableRecord = CType(Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead), DimStyleTableRecord)
                            LineAngularDimension2 = New LineAngularDimension2(Line1Start, Line1End, Line2Start, Line2End, ArcPoint, DimensionText, DimStyleTableRecord.ObjectId)
                            LineAngularDimension2.SetDatabaseDefaults()
                            If IsNothing(Layer) = False Then
                                LineAngularDimension2.Layer = Layer
                            End If
                            If UseUCS = True Then
                                LineAngularDimension2.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                            BlockTableRecord.AppendEntity(LineAngularDimension2)
                            Transaction.AddNewlyCreatedDBObject(LineAngularDimension2, True)
                            Transaction.Commit()
                            Return LineAngularDimension2
                        Else
                            Transaction.Abort()
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão angular
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="Line1Start">Ponto inicial da primeira linha</param>
        ''' <param name="Line1End">Ponto final da primeira linha</param>
        ''' <param name="Line2Start">Ponto inicial da outra linha</param>
        ''' <param name="Line2End">Ponto final da outra linha</param>
        ''' <param name="ArcPoint">Ponto de fixação</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>LineAngularDimension2</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawLineAngularDimension(Transaction As Transaction, DimStyleName As String, Line1Start As Point3d, Line1End As Point3d, Line2Start As Point3d, Line2End As Point3d, ByVal ArcPoint As Point3d, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As LineAngularDimension2
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim LineAngularDimension2 As LineAngularDimension2
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                If DimStyleTable.Has(DimStyleName) = True Then
                    DimStyleTableRecord = CType(Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead), DimStyleTableRecord)
                    LineAngularDimension2 = New LineAngularDimension2(Line1Start, Line1End, Line2Start, Line2End, ArcPoint, DimensionText, DimStyleTableRecord.ObjectId)
                    LineAngularDimension2.SetDatabaseDefaults()
                    If IsNothing(Layer) = False Then
                        LineAngularDimension2.Layer = Layer
                    End If
                    If UseUCS = True Then
                        LineAngularDimension2.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    BlockTableRecord.AppendEntity(LineAngularDimension2)
                    Transaction.AddNewlyCreatedDBObject(LineAngularDimension2, True)
                    Return LineAngularDimension2
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão ordenada
        ''' </summary>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="UseXaxis">Determina se usa eixo X</param>
        ''' <param name="DefiningPoint">Ponto de definição</param>
        ''' <param name="LeaderEndPoint">Ponto da seta</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>OrdinateDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawOrdinateDimension(ByVal DimStyleName As String, UseXaxis As Boolean, ByVal DefiningPoint As Point3d, ByVal LeaderEndPoint As Point3d, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As OrdinateDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim OrdinateDimension As OrdinateDimension
                Dim DimStr As String
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                        If DimStyleTable.Has(DimStyleName) = True Then
                            DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                            DimStr = String.Format("{0:f3}\P{1:f3}", DefiningPoint.X, DefiningPoint.Y)
                            OrdinateDimension = New OrdinateDimension(UseXaxis, DefiningPoint, LeaderEndPoint, DimStr, DimStyleTableRecord.ObjectId)
                            OrdinateDimension.SetDatabaseDefaults()
                            If IsNothing(Layer) = False Then
                                OrdinateDimension.Layer = Layer
                            End If
                            If UseUCS = True Then
                                OrdinateDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                            BlockTableRecord.AppendEntity(OrdinateDimension)
                            Transaction.AddNewlyCreatedDBObject(OrdinateDimension, True)
                            Transaction.Commit()
                            Return OrdinateDimension
                        Else
                            Transaction.Abort()
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão ordenada
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="UseXaxis">Determina se usa eixo X</param>
        ''' <param name="DefiningPoint">Ponto de definição</param>
        ''' <param name="LeaderEndPoint">Ponto da seta</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>OrdinateDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawOrdinateDimension(Transaction As Transaction, ByVal DimStyleName As String, UseXaxis As Boolean, ByVal DefiningPoint As Point3d, ByVal LeaderEndPoint As Point3d, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As OrdinateDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim OrdinateDimension As OrdinateDimension
                Dim DimStr As String
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                If DimStyleTable.Has(DimStyleName) = True Then
                    DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                    DimStr = String.Format("{0:f3}\P{1:f3}", DefiningPoint.X, DefiningPoint.Y)
                    OrdinateDimension = New OrdinateDimension(UseXaxis, DefiningPoint, LeaderEndPoint, DimStr, DimStyleTableRecord.ObjectId)
                    OrdinateDimension.SetDatabaseDefaults()
                    If IsNothing(Layer) = False Then
                        OrdinateDimension.Layer = Layer
                    End If
                    If UseUCS = True Then
                        OrdinateDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    BlockTableRecord.AppendEntity(OrdinateDimension)
                    Transaction.AddNewlyCreatedDBObject(OrdinateDimension, True)
                    Return OrdinateDimension
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão em arco
        ''' </summary>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="CenterPoint">Ponto central</param>
        ''' <param name="xLine1Point">Ponto inicial em X</param>
        ''' <param name="xLine2Point">Ponto final em X</param>
        ''' <param name="ArcPoint">Ponto do arco</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>ArcDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawArcDimension(ByVal DimStyleName As String, ByVal CenterPoint As Point3d, ByVal xLine1Point As Point3d, ByVal xLine2Point As Point3d, ByVal ArcPoint As Point3d, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As ArcDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim ArcDimension As ArcDimension
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                        If DimStyleTable.Has(DimStyleName) = True Then
                            DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                            ArcDimension = New ArcDimension(CenterPoint, xLine1Point, xLine2Point, ArcPoint, DimensionText, DimStyleTableRecord.ObjectId)
                            ArcDimension.SetDatabaseDefaults()
                            If IsNothing(Layer) = False Then
                                ArcDimension.Layer = Layer
                            End If
                            If UseUCS = True Then
                                ArcDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                            BlockTableRecord.AppendEntity(ArcDimension)
                            Transaction.AddNewlyCreatedDBObject(ArcDimension, True)
                            Transaction.Commit()
                            Return ArcDimension
                        Else
                            Transaction.Abort()
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão em arco
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="CenterPoint">Ponto central</param>
        ''' <param name="xLine1Point">Ponto inicial em X</param>
        ''' <param name="xLine2Point">Ponto final em X</param>
        ''' <param name="ArcPoint">Ponto do arco</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>ArcDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawArcDimension(Transaction As Transaction, ByVal DimStyleName As String, ByVal CenterPoint As Point3d, ByVal xLine1Point As Point3d, ByVal xLine2Point As Point3d, ByVal ArcPoint As Point3d, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As ArcDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim ArcDimension As ArcDimension
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                If DimStyleTable.Has(DimStyleName) = True Then
                    DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                    ArcDimension = New ArcDimension(CenterPoint, xLine1Point, xLine2Point, ArcPoint, DimensionText, DimStyleTableRecord.ObjectId)
                    ArcDimension.SetDatabaseDefaults()
                    If IsNothing(Layer) = False Then
                        ArcDimension.Layer = Layer
                    End If
                    If UseUCS = True Then
                        ArcDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    BlockTableRecord.AppendEntity(ArcDimension)
                    Transaction.AddNewlyCreatedDBObject(ArcDimension, True)
                    Return ArcDimension
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha a dimensão em diâmetro
        ''' </summary>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="ChordPoint">Ponto inicial de ancoragem</param>
        ''' <param name="FarChordPoint">Ponto final de ancoragem</param>
        ''' <param name="LeaderLength">Comprimento da seta</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>DiametricDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawDiametricDimension(ByVal DimStyleName As String, ByVal ChordPoint As Point3d, ByVal FarChordPoint As Point3d, ByVal LeaderLength As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As DiametricDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim DiametricDimension As DiametricDimension
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                        If DimStyleTable.Has(DimStyleName) = True Then
                            DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                            DiametricDimension = New DiametricDimension(ChordPoint, FarChordPoint, LeaderLength, DimensionText, DimStyleTableRecord.ObjectId)
                            DiametricDimension.SetDatabaseDefaults()
                            If IsNothing(Layer) = False Then
                                DiametricDimension.Layer = Layer
                            End If
                            If UseUCS = True Then
                                DiametricDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                            BlockTableRecord.AppendEntity(DiametricDimension)
                            Transaction.AddNewlyCreatedDBObject(DiametricDimension, True)
                            Transaction.Commit()
                            Return DiametricDimension
                        Else
                            Transaction.Abort()
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha a dimensão em diâmetro
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="ChordPoint">Ponto inicial de ancoragem</param>
        ''' <param name="FarChordPoint">Ponto final de ancoragem</param>
        ''' <param name="LeaderLength">Comprimento da seta</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>DiametricDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawDiametricDimension(Transaction As Transaction, ByVal DimStyleName As String, ByVal ChordPoint As Point3d, ByVal FarChordPoint As Point3d, ByVal LeaderLength As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As DiametricDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim DiametricDimension As DiametricDimension
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                If DimStyleTable.Has(DimStyleName) = True Then
                    DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                    DiametricDimension = New DiametricDimension(ChordPoint, FarChordPoint, LeaderLength, DimensionText, DimStyleTableRecord.ObjectId)
                    DiametricDimension.SetDatabaseDefaults()
                    If IsNothing(Layer) = False Then
                        DiametricDimension.Layer = Layer
                    End If
                    If UseUCS = True Then
                        DiametricDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    BlockTableRecord.AppendEntity(DiametricDimension)
                    Transaction.AddNewlyCreatedDBObject(DiametricDimension, True)
                    Return DiametricDimension
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão radial
        ''' </summary>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="CenterPoint">Ponto central</param>
        ''' <param name="ChordPoint">Ponto de ancoragem</param>
        ''' <param name="LeaderLength">Comprimento da seta</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>RadialDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawRadialDimension(ByVal DimStyleName As String, ByVal CenterPoint As Point3d, ByVal ChordPoint As Point3d, ByVal LeaderLength As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As RadialDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim RadialDimension As RadialDimension
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                        If DimStyleTable.Has(DimStyleName) = True Then
                            DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                            RadialDimension = New RadialDimension(CenterPoint, ChordPoint, LeaderLength, DimensionText, DimStyleTableRecord.ObjectId)
                            RadialDimension.SetDatabaseDefaults()
                            If IsNothing(Layer) = False Then
                                RadialDimension.Layer = Layer
                            End If
                            If UseUCS = True Then
                                RadialDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                            BlockTableRecord.AppendEntity(RadialDimension)
                            Transaction.AddNewlyCreatedDBObject(RadialDimension, True)
                            Transaction.Commit()
                            Return RadialDimension
                        Else
                            Transaction.Abort()
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão radial
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="CenterPoint">Ponto central</param>
        ''' <param name="ChordPoint">Ponto de ancoragem</param>
        ''' <param name="LeaderLength">Comprimento da seta</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>RadialDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawRadialDimension(Transaction As Transaction, ByVal DimStyleName As String, ByVal CenterPoint As Point3d, ByVal ChordPoint As Point3d, ByVal LeaderLength As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As RadialDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim RadialDimension As RadialDimension
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                If DimStyleTable.Has(DimStyleName) = True Then
                    DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                    RadialDimension = New RadialDimension(CenterPoint, ChordPoint, LeaderLength, DimensionText, DimStyleTableRecord.ObjectId)
                    RadialDimension.SetDatabaseDefaults()
                    If IsNothing(Layer) = False Then
                        RadialDimension.Layer = Layer
                    End If
                    If UseUCS = True Then
                        RadialDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    BlockTableRecord.AppendEntity(RadialDimension)
                    Transaction.AddNewlyCreatedDBObject(RadialDimension, True)
                    Return RadialDimension
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão radial
        ''' </summary>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="CenterPoint">Ponto central</param>
        ''' <param name="ChordPoint">Ponto de ancoragem</param>
        ''' <param name="OverridePoint">Ponto de substituição</param>
        ''' <param name="JogPoint">Ponto do caminho</param>
        ''' <param name="JogAngle">Caminho angular</param>
        ''' <param name="DimensionText">Text</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>RadialDimensionLarge</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawRadialDimensionLarge(ByVal DimStyleName As String, ByVal CenterPoint As Point3d, ByVal ChordPoint As Point3d, ByVal OverridePoint As Point3d, ByVal JogPoint As Point3d, ByVal JogAngle As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As RadialDimensionLarge
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim RadialDimensionLarge As RadialDimensionLarge
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                        If DimStyleTable.Has(DimStyleName) = True Then
                            DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                            RadialDimensionLarge = New RadialDimensionLarge(CenterPoint, ChordPoint, OverridePoint, JogPoint, JogAngle, DimensionText, DimStyleTableRecord.ObjectId)
                            RadialDimensionLarge.SetDatabaseDefaults()
                            If IsNothing(Layer) = False Then
                                RadialDimensionLarge.Layer = Layer
                            End If
                            If UseUCS = True Then
                                RadialDimensionLarge.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                            BlockTableRecord.AppendEntity(RadialDimensionLarge)
                            Transaction.AddNewlyCreatedDBObject(RadialDimensionLarge, True)
                            Transaction.Commit()
                            Return RadialDimensionLarge
                        Else
                            Transaction.Abort()
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão radial
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="CenterPoint">Ponto central</param>
        ''' <param name="ChordPoint">Ponto de ancoragem</param>
        ''' <param name="OverridePoint">Ponto de substituição</param>
        ''' <param name="JogPoint">Ponto do caminho</param>
        ''' <param name="JogAngle">Caminho angular</param>
        ''' <param name="DimensionText">Text</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>RadianDimensionLarge</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawRadialDimensionLarge(Transaction As Transaction, ByVal DimStyleName As String, ByVal CenterPoint As Point3d, ByVal ChordPoint As Point3d, ByVal OverridePoint As Point3d, ByVal JogPoint As Point3d, ByVal JogAngle As Double, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As RadialDimensionLarge
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim RadialDimensionLarge As RadialDimensionLarge
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                If DimStyleTable.Has(DimStyleName) = True Then
                    DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                    RadialDimensionLarge = New RadialDimensionLarge(CenterPoint, ChordPoint, OverridePoint, JogPoint, JogAngle, DimensionText, DimStyleTableRecord.ObjectId)
                    RadialDimensionLarge.SetDatabaseDefaults()
                    If IsNothing(Layer) = False Then
                        RadialDimensionLarge.Layer = Layer
                    End If
                    If UseUCS = True Then
                        RadialDimensionLarge.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    BlockTableRecord.AppendEntity(RadialDimensionLarge)
                    Transaction.AddNewlyCreatedDBObject(RadialDimensionLarge, True)
                    Return RadialDimensionLarge
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão angular em 3 pontos
        ''' </summary>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="CenterPoint">Ponto central</param>
        ''' <param name="Line1Point">Ponto na linha 1</param>
        ''' <param name="Line2Point">Ponto na linha 2</param>
        ''' <param name="ArcPoint">Ponto do arco</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Point3AngularDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawPoint3AngularDimension(ByVal DimStyleName As String, ByVal CenterPoint As Point3d, ByVal Line1Point As Point3d, ByVal Line2Point As Point3d, ByVal ArcPoint As Point3d, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Point3AngularDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim Point3AngularDimension As Point3AngularDimension
                Using Editor.Document.LockDocument
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                        If DimStyleTable.Has(DimStyleName) = True Then
                            DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                            Point3AngularDimension = New Point3AngularDimension(CenterPoint, Line1Point, Line2Point, ArcPoint, DimensionText, DimStyleTableRecord.ObjectId)
                            Point3AngularDimension.SetDatabaseDefaults()
                            If IsNothing(Layer) = False Then
                                Point3AngularDimension.Layer = Layer
                            End If
                            If UseUCS = True Then
                                Point3AngularDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                            BlockTableRecord.AppendEntity(Point3AngularDimension)
                            Transaction.AddNewlyCreatedDBObject(Point3AngularDimension, True)
                            Transaction.Commit()
                            Return Point3AngularDimension
                        Else
                            Transaction.Abort()
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha dimensão angular em 3 pontos
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="DimStyleName">Estilo de dimensão</param>
        ''' <param name="CenterPoint">Ponto central</param>
        ''' <param name="Line1Point">Ponto na linha 1</param>
        ''' <param name="Line2Point">Ponto na linha 2</param>
        ''' <param name="ArcPoint">Ponto do arco</param>
        ''' <param name="DimensionText">Texto</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Point3AngularDimension</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawPoint3AngularDimension(Transaction As Transaction, ByVal DimStyleName As String, ByVal CenterPoint As Point3d, ByVal Line1Point As Point3d, ByVal Line2Point As Point3d, ByVal ArcPoint As Point3d, Optional DimensionText As String = "", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Point3AngularDimension
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Editor As Editor = Document.Editor
                Dim Database As Database = Document.Database
                Dim BlockTableRecord As BlockTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim Point3AngularDimension As Point3AngularDimension
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                DimStyleTable = Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead)
                If DimStyleTable.Has(DimStyleName) = True Then
                    DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                    Point3AngularDimension = New Point3AngularDimension(CenterPoint, Line1Point, Line2Point, ArcPoint, DimensionText, DimStyleTableRecord.ObjectId)
                    Point3AngularDimension.SetDatabaseDefaults()
                    If IsNothing(Layer) = False Then
                        Point3AngularDimension.Layer = Layer
                    End If
                    If UseUCS = True Then
                        Point3AngularDimension.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    BlockTableRecord.AppendEntity(Point3AngularDimension)
                    Transaction.AddNewlyCreatedDBObject(Point3AngularDimension, True)
                    Return Point3AngularDimension
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha a polyline
        ''' </summary>
        ''' <param name="Point3dCollection">Coleção de pontos 3d</param>
        ''' <param name="Closed">Determina se a polyline será fechada</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Polyline</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawPolyline(Point3dCollection As Point3dCollection, Closed As Boolean, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Polyline
            Dim Polyline As Polyline
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim BlockTableRecord As BlockTableRecord
            Dim Point2d As Point2d
            Dim Point2dCollection As New Point2dCollection
            Dim Normal As Vector3d = New Vector3d(0, 0, 1)
            Normal = Normal.TransformBy(Editor.CurrentUserCoordinateSystem)
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    Try
                        For Each Point3d As Point3d In Point3dCollection
                            Point2d = Point3d.Convert2d(New Plane(Point3d.Origin, Vector3d.ZAxis))
                            Point2dCollection.Add(Point2d)
                        Next
                        Polyline = New Polyline(Point2dCollection.Count)
                        With Polyline
                            .Normal = Normal
                            If IsNothing(Layer) = False Then
                                .Layer = Layer
                            End If
                            For Index As Integer = 0 To Point2dCollection.Count - 1
                                Point2d = Point2dCollection.Item(Index)
                                .AddVertexAt(Index, Point2d, 0, 0, 0)
                            Next
                            .Closed = Closed
                            If UseUCS = True Then
                                .TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
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
        ''' Desenha a polyline
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Point3dCollection">Coleção de pontos 3d</param>
        ''' <param name="Closed">Determina se a polyline será fechada</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Polyline</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawPolyline(Transaction As Transaction, Point3dCollection As Point3dCollection, Closed As Boolean, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Polyline
            Dim Polyline As Polyline
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim BlockTableRecord As BlockTableRecord
            Dim Point2d As Point2d
            Dim Point2dCollection As New Point2dCollection
            Dim Normal As Vector3d = New Vector3d(0, 0, 1)
            Normal = Normal.TransformBy(Editor.CurrentUserCoordinateSystem)
            Try
                For Each Point3d As Point3d In Point3dCollection
                    Point2d = Point3d.Convert2d(New Plane(Point3d.Origin, Vector3d.ZAxis))
                    Point2dCollection.Add(Point2d)
                Next
                Polyline = New Polyline(Point2dCollection.Count)
                With Polyline
                    .Normal = Normal
                    If IsNothing(Layer) = False Then
                        .Layer = Layer
                    End If
                    For Index As Integer = 0 To Point2dCollection.Count - 1
                        Point2d = Point2dCollection.Item(Index)
                        .AddVertexAt(Index, Point2d, 0, 0, 0)
                    Next
                    .Closed = Closed
                    If UseUCS = True Then
                        .TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                End With
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                BlockTableRecord.AppendEntity(Polyline)
                Transaction.AddNewlyCreatedDBObject(Polyline, True)
                Return Polyline
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha OffSet
        ''' </summary>
        ''' <param name="Entity">Entidade de origem</param>
        ''' <param name="OffSetValue">Valor do OffSet</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>DBObjectCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawOffSet(Entity As Entity, OffSetValue As Double, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As DBObjectCollection
            DrawOffSet = New DBObjectCollection
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim BlockTableRecord As BlockTableRecord
            Dim Curve As Curve
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    Curve = Entity
                    DrawOffSet = Curve.GetOffsetCurves(OffSetValue)
                    BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                    For Each EntityResult As Entity In DrawOffSet
                        If IsNothing(Layer) = False Then
                            EntityResult.Layer = Layer
                        End If
                        If UseUCS = True Then
                            EntityResult.TransformBy(Editor.CurrentUserCoordinateSystem)
                        End If
                        BlockTableRecord.AppendEntity(EntityResult)
                        Transaction.AddNewlyCreatedDBObject(EntityResult, True)
                    Next
                    Transaction.Commit()
                End Using
            End Using
            If DrawOffSet.Count = 0 Then
                Return Nothing
            Else
                Return DrawOffSet
            End If
        End Function

        ''' <summary>
        ''' Desenha OffSet
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Entity">Entidade de origem</param>
        ''' <param name="OffSetValue">Valor do OffSet</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>DBObjectCollection</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawOffSet(Transaction As Transaction, Entity As Entity, OffSetValue As Double, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As DBObjectCollection
            DrawOffSet = New DBObjectCollection
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim BlockTableRecord As BlockTableRecord
            Dim Curve As Curve
            Curve = Entity
            DrawOffSet = Curve.GetOffsetCurves(OffSetValue)
            BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
            For Each EntityResult As Entity In DrawOffSet
                If IsNothing(Layer) = False Then
                    EntityResult.Layer = Layer
                End If
                If UseUCS = True Then
                    EntityResult.TransformBy(Editor.CurrentUserCoordinateSystem)
                End If
                BlockTableRecord.AppendEntity(EntityResult)
                Transaction.AddNewlyCreatedDBObject(EntityResult, True)
            Next
            If DrawOffSet.Count = 0 Then
                Return Nothing
            Else
                Return DrawOffSet
            End If
        End Function

        ''' <summary>
        ''' Desenha Multiline
        ''' </summary>
        ''' <param name="Point3dCollection">Coleção de vértices</param>
        ''' <param name="Justification">Justificação</param>
        ''' <param name="Scale">Escala</param>
        ''' <param name="StyleName">Nome do estilo</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Mline</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawMultiLine(Point3dCollection As Point3dCollection, Justification As Autodesk.AutoCAD.DatabaseServices.MlineJustification, Scale As Double, Optional StyleName As String = "STANDARD", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Mline
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim MlineStyleID As ObjectId
            Dim MlineStyle As MlineStyle
            Dim DBDictionary As DBDictionary
            Dim Mline As Mline
            Dim BlockTableRecord As BlockTableRecord
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    Try
                        DBDictionary = Transaction.GetObject(Database.MLStyleDictionaryId, OpenMode.ForRead)
                        For Each DictionaryEntry As DictionaryEntry In DBDictionary
                            MlineStyle = Transaction.GetObject(DictionaryEntry.Value, OpenMode.ForRead)
                            If MlineStyle.Name = StyleName Then
                                MlineStyleID = MlineStyle.ObjectId
                                Exit For
                            End If
                        Next
                        Mline = New Mline
                        With Mline
                            If IsNothing(Layer) = False Then
                                .Layer = Layer
                            End If
                            .Style = MlineStyleID
                            .Normal = Vector3d.ZAxis
                            .Justification = Justification
                            .Scale = Scale
                            For Each Point3d As Point3d In Point3dCollection
                                .AppendSegment(Point3d)
                            Next
                            If UseUCS = True Then
                                .TransformBy(Editor.CurrentUserCoordinateSystem)
                            End If
                        End With
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                        BlockTableRecord.AppendEntity(Mline)
                        Transaction.AddNewlyCreatedDBObject(Mline, True)
                        Transaction.Commit()
                        Return Mline
                    Catch
                        Transaction.Abort()
                        Return Nothing
                    End Try
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Desenha Multiline
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Point3dCollection">Coleção de vértices</param>
        ''' <param name="Justification">Justificação</param>
        ''' <param name="Scale">Escala</param>
        ''' <param name="StyleName">Nome do estilo</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Mline</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawMultiLine(Transaction As Transaction, Point3dCollection As Point3dCollection, Justification As Autodesk.AutoCAD.DatabaseServices.MlineJustification, Scale As Double, Optional StyleName As String = "STANDARD", Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Mline
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim MlineStyleID As ObjectId
            Dim MlineStyle As MlineStyle
            Dim DBDictionary As DBDictionary
            Dim Mline As Mline
            Dim BlockTableRecord As BlockTableRecord
            Try
                DBDictionary = Transaction.GetObject(Database.MLStyleDictionaryId, OpenMode.ForRead)
                For Each DictionaryEntry As DictionaryEntry In DBDictionary
                    MlineStyle = Transaction.GetObject(DictionaryEntry.Value, OpenMode.ForRead)
                    If MlineStyle.Name = StyleName Then
                        MlineStyleID = MlineStyle.ObjectId
                        Exit For
                    End If
                Next
                Mline = New Mline
                With Mline
                    If IsNothing(Layer) = False Then
                        .Layer = Layer
                    End If
                    .Style = MlineStyleID
                    .Normal = Vector3d.ZAxis
                    .Justification = Justification
                    .Scale = Scale
                    For Each Point3d As Point3d In Point3dCollection
                        .AppendSegment(Point3d)
                    Next
                    If UseUCS = True Then
                        .TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                End With
                BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                BlockTableRecord.AppendEntity(Mline)
                Transaction.AddNewlyCreatedDBObject(Mline, True)
                Return Mline
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Desenha texto
        ''' </summary>
        ''' <param name="Contents">Texto</param>
        ''' <param name="Location">Local de inserção</param>
        ''' <param name="TextHeight">Altura do texto</param>
        ''' <param name="TextStyle">Estilo de texto</param>
        ''' <param name="Attachment">Posicionamento</param>
        ''' <Rotation>Rotação (radianos)</Rotation>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>MText</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawMtext(Contents As String, Location As Point3d, TextHeight As Double, Optional TextStyle As String = "Standard", Optional Attachment As Autodesk.AutoCAD.DatabaseServices.AttachmentPoint = AttachmentPoint.MiddleCenter, Optional Rotation As Double = 0, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As MText
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim MText As MText
            Dim BlockTableRecord As BlockTableRecord
            Dim TextStyleTable As TextStyleTable
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                    TextStyleTable = Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForRead)
                    MText = New MText
                    With MText
                        If IsNothing(Layer) = False Then
                            .Layer = Layer
                        End If
                        .TextStyleId = TextStyleTable(TextStyle)
                        .Location = Location
                        .Contents = Contents
                        .TextHeight = TextHeight
                        .Attachment = Attachment
                        .Rotation = Rotation
                        If UseUCS = True Then
                            .TransformBy(Editor.CurrentUserCoordinateSystem)
                        End If
                    End With
                    BlockTableRecord.AppendEntity(MText)
                    Transaction.AddNewlyCreatedDBObject(MText, True)
                    Transaction.Commit()
                End Using
            End Using
            Return MText
        End Function

        ''' <summary>
        ''' Desenha texto
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Contents">Texto</param>
        ''' <param name="Location">Local de inserção</param>
        ''' <param name="TextHeight">Altura do texto</param>
        ''' <param name="TextStyle">Estilo de texto</param>
        ''' <param name="Attachment">Posicionamento</param>
        '''<param name="Rotation">Rotação (radianos)</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>MText</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawMtext(Transaction As Transaction, Contents As String, Location As Point3d, TextHeight As Double, Optional TextStyle As String = "Standard", Optional Attachment As Autodesk.AutoCAD.DatabaseServices.AttachmentPoint = AttachmentPoint.MiddleCenter, Optional Rotation As Double = 0, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As MText
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim MText As MText
            Dim BlockTableRecord As BlockTableRecord
            Dim TextStyleTable As TextStyleTable
            BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
            TextStyleTable = Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForRead)
            MText = New MText
            With MText
                If IsNothing(Layer) = False Then
                    .Layer = Layer
                End If
                .TextStyleId = TextStyleTable(TextStyle)
                .Location = Location
                .Contents = Contents
                .TextHeight = TextHeight
                .Attachment = Attachment
                .Rotation = Rotation
                If UseUCS = True Then
                    .TransformBy(Editor.CurrentUserCoordinateSystem)
                End If
            End With
            BlockTableRecord.AppendEntity(MText)
            Transaction.AddNewlyCreatedDBObject(MText, True)
            Return MText
        End Function

        ''' <summary>
        ''' Desenha DBText
        ''' </summary>
        ''' <param name="TextString">Texto</param>
        ''' <param name="Position">Local da inserção</param>
        ''' <param name="Height">Altura do texto</param>
        ''' <param name="TextStyle">Estilo de texto</param>
        ''' <param name="Justify">Posicionamento</param>
        ''' <param name="Rotation">Rotação (radianos)</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>DBText</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawDBText(TextString As String, Position As Point3d, Height As Double, Optional TextStyle As String = "Standard", Optional Justify As Autodesk.AutoCAD.DatabaseServices.AttachmentPoint = AttachmentPoint.MiddleCenter, Optional Rotation As Double = 0, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As DBText
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTableRecord As BlockTableRecord
            Dim DBText As DBText = Nothing
            Dim TextStyleTable As TextStyleTable
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                    TextStyleTable = Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForRead)
                    DBText = New DBText
                    With DBText
                        .TextStyleId = TextStyleTable(TextStyle)
                        .TextString = TextString
                        .Height = Height
                        .HorizontalMode = TextHorizontalMode.TextCenter
                        .VerticalMode = TextVerticalMode.TextVerticalMid
                        .AlignmentPoint = Position
                        .Justify = Justify
                        .Rotation = Rotation
                        .Position = Position
                        If IsNothing(Layer) = False Then
                            .Layer = Layer
                        End If
                        If UseUCS = True Then
                            .TransformBy(Editor.CurrentUserCoordinateSystem)
                        End If
                    End With
                    BlockTableRecord.AppendEntity(DBText)
                    Transaction.AddNewlyCreatedDBObject(DBText, True)
                    Transaction.Commit()
                End Using
            End Using
            Return DBText
        End Function

        ''' <summary>
        ''' Desenha DBText
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="TextString">Texto</param>
        ''' <param name="Position">Local da inserção</param>
        ''' <param name="Height">Altura do texto</param>
        ''' <param name="TextStyle">Estilo de texto</param>
        ''' <param name="Justify">Posicionamento</param>
        ''' <param name="Rotation">Rotação (radianos)</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>DBText</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawDBText(Transaction As Transaction, TextString As String, Position As Point3d, Height As Double, Optional TextStyle As String = "Standard", Optional Justify As Autodesk.AutoCAD.DatabaseServices.AttachmentPoint = AttachmentPoint.MiddleCenter, Optional Rotation As Double = 0, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As DBText
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTableRecord As BlockTableRecord
            Dim DBText As DBText = Nothing
            Dim TextStyleTable As TextStyleTable
            BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
            TextStyleTable = Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForRead)
            DBText = New DBText
            With DBText
                .TextStyleId = TextStyleTable(TextStyle)
                .TextString = TextString
                .Height = Height
                .HorizontalMode = TextHorizontalMode.TextCenter
                .VerticalMode = TextVerticalMode.TextVerticalMid
                .AlignmentPoint = Position
                .Justify = Justify
                .Rotation = Rotation
                .Position = Position
                If IsNothing(Layer) = False Then
                    .Layer = Layer
                End If
                If UseUCS = True Then
                    .TransformBy(Editor.CurrentUserCoordinateSystem)
                End If
            End With
            BlockTableRecord.AppendEntity(DBText)
            Transaction.AddNewlyCreatedDBObject(DBText, True)
            Return DBText
        End Function

        ''' <summary>
        ''' Desenha WipeOut
        ''' </summary>
        ''' <param name="Center">Posição</param>
        ''' <param name="Reach">Alcance a partir do centro</param>
        ''' <param name="NumberOfVertices">Número do vértices</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Wipeout</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawWipeout(Center As Point3d, Reach As Double, Optional NumberOfVertices As Integer = 12, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Wipeout
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTableRecord As BlockTableRecord
            Dim Wipeout As Wipeout
            Dim Point2dCollection As New Point2dCollection
            Dim Curve As Curve
            Dim Point3d As Point3d
            Dim Lenght As Double
            Dim [Step] As Double
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite, False)
                    Curve = New Circle(Center, Vector3d.ZAxis, Reach)
                    Lenght = Curve.GetDistanceAtParameter(Curve.EndParam) - Curve.GetDistanceAtParameter(Curve.StartParam)
                    [Step] = (Lenght / NumberOfVertices)
                    For CurvePosition As Double = 0 To Lenght Step [Step]
                        Point3d = Curve.GetPointAtDist(CurvePosition)
                        Point2dCollection.Add(Engine2.Geometry.Point3dToPoint2d(Curve.GetClosestPointTo(Point3d, False)))
                    Next
                    If Point2dCollection.Item(0).Equals(Point2dCollection.Item(Point2dCollection.Count - 1)) = False Then
                        Point2dCollection.Add(Point2dCollection.Item(0))
                    End If
                    Curve.Dispose()
                    Wipeout = New Wipeout()
                    If IsNothing(Layer) = False Then
                        Wipeout.Layer = Layer
                    End If
                    Wipeout.SetDatabaseDefaults(Database)
                    Wipeout.SetFrom(Point2dCollection, New Vector3d(0.0, 0.0, 0.1))
                    If UseUCS = True Then
                        Wipeout.TransformBy(Editor.CurrentUserCoordinateSystem)
                    End If
                    BlockTableRecord.AppendEntity(Wipeout)
                    Transaction.AddNewlyCreatedDBObject(Wipeout, True)
                    Transaction.Commit()
                End Using
            End Using
            Return Wipeout
        End Function

        ''' <summary>
        ''' Desenha WipeOut
        ''' </summary>
        ''' <param name="Center">Posição</param>
        ''' <param name="Reach">Alcance a partir do centro</param>
        ''' <param name="NumberOfVertices">Número do vértices</param>
        ''' <param name="Layer">Camada</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <returns>Wipeout</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawWipeout(Transaction As Transaction, Center As Point3d, Reach As Double, Optional NumberOfVertices As Integer = 12, Optional Layer As Object = Nothing, Optional UseUCS As Boolean = True) As Wipeout
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim BlockTableRecord As BlockTableRecord
            Dim Wipeout As Wipeout
            Dim Point2dCollection As New Point2dCollection
            Dim Curve As Curve
            Dim Point3d As Point3d
            Dim Lenght As Double
            Dim [Step] As Double
            BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite, False)
            Curve = New Circle(Center, Vector3d.ZAxis, Reach)
            Lenght = Curve.GetDistanceAtParameter(Curve.EndParam) - Curve.GetDistanceAtParameter(Curve.StartParam)
            [Step] = (Lenght / NumberOfVertices)
            For CurvePosition As Double = 0 To Lenght Step [Step]
                Point3d = Curve.GetPointAtDist(CurvePosition)
                Point2dCollection.Add(Engine2.Geometry.Point3dToPoint2d(Curve.GetClosestPointTo(Point3d, False)))
            Next
            If Point2dCollection.Item(0).Equals(Point2dCollection.Item(Point2dCollection.Count - 1)) = False Then
                Point2dCollection.Add(Point2dCollection.Item(0))
            End If
            Curve.Dispose()
            Wipeout = New Wipeout()
            If IsNothing(Layer) = False Then
                Wipeout.Layer = Layer
            End If
            Wipeout.SetDatabaseDefaults(Database)
            Wipeout.SetFrom(Point2dCollection, New Vector3d(0.0, 0.0, 0.1))
            If UseUCS = True Then
                Wipeout.TransformBy(Editor.CurrentUserCoordinateSystem)
            End If
            BlockTableRecord.AppendEntity(Wipeout)
            Transaction.AddNewlyCreatedDBObject(Wipeout, True)
            Return Wipeout
        End Function

        ''' <summary>
        ''' Desenha espelhamento de entidades
        ''' </summary>
        ''' <param name="Entity">Entidade a ser espelhada</param>
        ''' <param name="LineSegment3d">Segmento de linha que corresponde a mediatrix</param>
        ''' <param name="Erased">Determina se a entidade original será excluída</param>
        ''' <returns>Entity\Nothing</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawMirror(Entity As Entity, LineSegment3d As LineSegment3d, Erased As Boolean) As Entity
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim Clone As Entity = Nothing
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            Dim Matrix3d As Matrix3d
            Dim DBText As DBText
            Dim Line3d As Line3d = New Line3d(LineSegment3d.StartPoint, LineSegment3d.EndPoint)
            Dim Position As Point3d
            Dim Rotation As Double
            Dim MidAngle As Double
            Using Editor.Document.LockDocument
                Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                    Try
                        BlockTable = Transaction.GetObject(Document.Database.BlockTableId, OpenMode.ForRead)
                        BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite, False)
                        Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
                        Clone = Entity.Clone
                        Select Case Clone.GetType.Name
                            Case "DBText"
                                DBText = Clone
                                Position = Engine2.Geometry.MirrorPoint(DBText.Position, LineSegment3d)
                                With DBText
                                    MidAngle = Engine2.Geometry.GetAngle(LineSegment3d.StartPoint, LineSegment3d.EndPoint)
                                    Rotation = Engine2.Geometry.RadianToDegree(.Rotation)
                                    Rotation = Engine2.Geometry.MirrorAngle(Rotation, MidAngle)
                                    Rotation = Engine2.Geometry.DegreeToRadian(Rotation)
                                    .Rotation = Rotation
                                    Position = Engine2.Geometry.MirrorPoint(.AlignmentPoint, LineSegment3d)
                                    .AlignmentPoint = Position
                                    .Position = Position
                                    Select Case .Justify
                                        Case AttachmentPoint.BaseLeft
                                            .Justify = AttachmentPoint.BaseRight
                                        Case AttachmentPoint.BaseRight
                                            .Justify = AttachmentPoint.BaseLeft
                                        Case AttachmentPoint.BottomLeft
                                            .Justify = AttachmentPoint.BottomRight
                                        Case AttachmentPoint.BottomRight
                                            .Justify = AttachmentPoint.BottomLeft
                                        Case AttachmentPoint.MiddleLeft
                                            .Justify = AttachmentPoint.MiddleRight
                                        Case AttachmentPoint.MiddleRight
                                            .Justify = AttachmentPoint.MiddleLeft
                                        Case AttachmentPoint.TopLeft
                                            .Justify = AttachmentPoint.TopRight
                                        Case AttachmentPoint.TopRight
                                            .Justify = AttachmentPoint.TopLeft
                                        Case Else
                                            .Justify = .Justify
                                    End Select
                                End With
                                Clone = DBText
                            Case Else
                                Matrix3d = Matrix3d.Mirroring(Line3d)
                                Clone.TransformBy(Matrix3d)
                        End Select
                        BlockTableRecord.AppendEntity(Clone)
                        Transaction.AddNewlyCreatedDBObject(Clone, True)
                        If Erased = True Then
                            Entity.Erase()
                        End If
                        If Erased = True Then
                            Entity.Erase()
                        End If
                        Transaction.Commit()
                    Catch
                        Transaction.Abort()
                        Clone = Nothing
                    End Try
                End Using
            End Using
            Return Clone
        End Function

        ''' <summary>
        ''' Desenha espelhamento de entidades
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Entity">Entidade a ser espelhada</param>
        ''' <param name="LineSegment3d">Segmento de linha que corresponde a mediatrix</param>
        ''' <param name="Erased">Determina se a entidade original será excluída</param>
        ''' <returns>Entity\Nothing</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawMirror(Transaction As Transaction, Entity As Entity, LineSegment3d As LineSegment3d, Erased As Boolean) As Entity
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Database As Database = Document.Database
            Dim Editor As Editor = Document.Editor
            Dim Clone As Entity = Nothing
            Dim BlockTable As BlockTable
            Dim BlockTableRecord As BlockTableRecord
            Dim Matrix3d As Matrix3d
            Dim DBText As DBText
            Dim Line3d As Line3d = New Line3d(LineSegment3d.StartPoint, LineSegment3d.EndPoint)
            Dim Position As Point3d
            Dim Rotation As Double
            Dim MidAngle As Double
            BlockTable = Transaction.GetObject(Document.Database.BlockTableId, OpenMode.ForRead)
            BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite, False)
            Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead)
            Clone = Entity.Clone
            Select Case Clone.GetType.Name
                Case "DBText"
                    DBText = Clone
                    Position = Engine2.Geometry.MirrorPoint(DBText.Position, LineSegment3d)
                    With DBText
                        MidAngle = Engine2.Geometry.GetAngle(LineSegment3d.StartPoint, LineSegment3d.EndPoint)
                        Rotation = Engine2.Geometry.RadianToDegree(.Rotation)
                        Rotation = Engine2.Geometry.MirrorAngle(Rotation, MidAngle)
                        Rotation = Engine2.Geometry.DegreeToRadian(Rotation)
                        .Rotation = Rotation
                        Position = Engine2.Geometry.MirrorPoint(.AlignmentPoint, LineSegment3d)
                        .AlignmentPoint = Position
                        .Position = Position
                        Select Case .Justify
                            Case AttachmentPoint.BaseLeft
                                .Justify = AttachmentPoint.BaseRight
                            Case AttachmentPoint.BaseRight
                                .Justify = AttachmentPoint.BaseLeft
                            Case AttachmentPoint.BottomLeft
                                .Justify = AttachmentPoint.BottomRight
                            Case AttachmentPoint.BottomRight
                                .Justify = AttachmentPoint.BottomLeft
                            Case AttachmentPoint.MiddleLeft
                                .Justify = AttachmentPoint.MiddleRight
                            Case AttachmentPoint.MiddleRight
                                .Justify = AttachmentPoint.MiddleLeft
                            Case AttachmentPoint.TopLeft
                                .Justify = AttachmentPoint.TopRight
                            Case AttachmentPoint.TopRight
                                .Justify = AttachmentPoint.TopLeft
                            Case Else
                                .Justify = .Justify
                        End Select
                    End With
                    Clone = DBText
                Case Else
                    Matrix3d = Matrix3d.Mirroring(Line3d)
                    Clone.TransformBy(Matrix3d)
            End Select
            BlockTableRecord.AppendEntity(Clone)
            Transaction.AddNewlyCreatedDBObject(Clone, True)
            If Erased = True Then
                Entity.Erase()
            End If
            Return Clone
        End Function

        ''' <summary>
        ''' Copia entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <param name="PointDisplacement">Ponto de deslocamento</param>
        ''' <returns>Entity</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawCopy(Entity As Entity, PointReference As Point3d, PointDisplacement As Point3d) As Entity
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim BlockTableRecord As BlockTableRecord
            Dim Matrix3d As Matrix3d = Matrix3d.Displacement(PointReference.GetVectorTo(PointDisplacement))
            Using DocumentLock As DocumentLock = Document.LockDocument()
                Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                    BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
                    Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead, False)
                    Entity = Entity.Clone
                    Entity.TransformBy(Matrix3d)
                    BlockTableRecord.AppendEntity(Entity)
                    Transaction.AddNewlyCreatedDBObject(Entity, True)
                    Transaction.Commit()
                End Using
            End Using
            Return Entity
        End Function

        ''' <summary>
        ''' Copia entidade
        ''' </summary>
        ''' <param name="Transaction">Transação</param>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="PointReference">Ponto de referência</param>
        ''' <param name="PointDisplacement">Ponto de deslocamento</param>
        ''' <returns>Entity</returns>
        ''' <remarks></remarks>
        Public Shared Function DrawCopy(Transaction As Transaction, Entity As Entity, PointReference As Point3d, PointDisplacement As Point3d) As Entity
            Dim Document As Document = Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Document.Editor
            Dim Database As Database = Document.Database
            Dim BlockTableRecord As BlockTableRecord
            Dim Matrix3d As Matrix3d = Matrix3d.Displacement(PointReference.GetVectorTo(PointDisplacement))
            BlockTableRecord = Transaction.GetObject(Database.CurrentSpaceId, OpenMode.ForWrite)
            Entity = Transaction.GetObject(Entity.ObjectId, OpenMode.ForRead, False)
            Entity = Entity.Clone
            Entity.TransformBy(Matrix3d)
            BlockTableRecord.AppendEntity(Entity)
            Transaction.AddNewlyCreatedDBObject(Entity, True)
            Return Entity
        End Function

        ''' <summary>
        ''' Adiciona hachura na entidade
        ''' </summary>
        ''' <param name="Entity">Entidade</param>
        ''' <param name="HatchName">Nome da hachura</param>
        ''' <param name="HatchScale">Escala</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DrawHatch(ByVal Entity As Entity, ByVal HatchName As String, ByVal HatchScale As Double, HatchAngle As Double, Optional Transparency As Byte = 50) As Hatch
            Try
                Dim Hatch As Hatch = Nothing
                Dim acadDocument As Document = Application.DocumentManager.MdiActiveDocument
                Dim acadDatabase As Database = acadDocument.Database
                Using Transaction As Transaction = acadDatabase.TransactionManager.StartTransaction()
                    Dim BlockTable As BlockTable = Transaction.GetObject(acadDatabase.BlockTableId, OpenMode.ForRead)
                    Dim BlockTableRecord As BlockTableRecord = Transaction.GetObject(BlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim ObjectIdCollection As ObjectIdCollection = New ObjectIdCollection
                    ObjectIdCollection.Add(Entity.ObjectId)
                    Hatch = New Hatch()
                    BlockTableRecord.AppendEntity(Hatch)
                    Transaction.AddNewlyCreatedDBObject(Hatch, True)
                    With Hatch
                        .SetDatabaseDefaults()
                        .SetHatchPattern(HatchPatternType.PreDefined, HatchName)
                        If HatchName = "SOLID" Then
                            .Transparency = New Autodesk.AutoCAD.Colors.Transparency(Transparency)
                        End If
                        .Associative = True
                        .PatternScale = HatchScale
                        .PatternAngle = HatchAngle
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
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DrawHatch(Transaction As Transaction, ByVal Entity As Entity, ByVal HatchName As String, ByVal HatchScale As Double, HatchAngle As Double, Optional Transparency As Byte = 50) As Hatch
            Try
                Dim Hatch As Hatch = Nothing
                Dim acadDocument As Document = Application.DocumentManager.MdiActiveDocument
                Dim acadDatabase As Database = acadDocument.Database
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
                    If HatchName = "SOLID" Then
                        .Transparency = New Autodesk.AutoCAD.Colors.Transparency(Transparency)
                    End If
                    .Associative = True
                    .PatternScale = HatchScale
                    .PatternAngle = HatchAngle
                    .AppendLoop(HatchLoopTypes.Outermost, ObjectIdCollection)
                    .EvaluateHatch(True)
                End With
                Return Hatch
            Catch
                Return Nothing
            End Try
        End Function

#Region "DESENHO DE VETORES"

        ''' <summary>
        ''' Enumera os tipos de setas
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eLeaderStyle

            ''' <summary>
            ''' Seta nas 2 direções
            ''' </summary>
            ''' <remarks></remarks>
            DoubleLeader = 0

            ''' <summary>
            ''' Seta no ponto inicial
            ''' </summary>
            ''' <remarks></remarks>
            StartPointLeader = 1

            ''' <summary>
            ''' Seta no ponto final
            ''' </summary>
            ''' <remarks></remarks>
            EndPointLeader = 2

        End Enum

        ''' <summary>
        ''' Desenha vetor em formato de seta
        ''' </summary>
        ''' <param name="LeaderStyle">Estilo de seta</param>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="EndPoint">Ponto final</param>
        ''' <param name="ColorIndex">Cor</param>
        ''' <param name="Highlight">Determina se o desenho será destacado</param>
        ''' <param name="Refresh">Determina se os vetores existentes serão excluídos durante o novo desenho</param>
        ''' <param name="LeaderSize">Tamanho da seta (10 = 1/10 do comprimento)</param>
        ''' <remarks></remarks>
        Public Shared Sub DrawLeaderVector(LeaderStyle As eLeaderStyle, StartPoint As Point3d, EndPoint As Point3d, ColorIndex As Integer, Optional Highlight As Boolean = False, Optional Refresh As Boolean = True, Optional LeaderSize As Double = 10)

            'Declarações
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Distance As Double
            Dim Lenght As Double
            Dim Angle As Double
            Dim Points As New List(Of Point3dCollection)
            Dim LeaderAngle As Double
            Dim TempPoint As Point3d
            Dim TempPoint2 As Point3d
            Dim TempAngle As Double
            Dim TempDist As Double

            'Limpa os vetores anteriores
            If Refresh = True Then
                'Editor.Regen()
                Engine2.EntityInteration.Redraw2()
            End If

            'Obtem o comprimento do vetor
            Lenght = StartPoint.DistanceTo(EndPoint)

            'Obtem o ângulo do vetor
            Angle = Engine2.Geometry.GetAngle(StartPoint, EndPoint, Geometry.eAngleFormat.Degrees)

            'Obtem o tamanho da seta
            LeaderSize = (Lenght / LeaderSize)

            'Distância ângular da seta
            TempDist = LeaderSize
            Distance = Engine2.Geometry.Hypotenuse(TempDist, TempDist)

            'Determina o estilo da seta
            Select Case LeaderStyle
                Case eLeaderStyle.DoubleLeader

                    'Adiciona o 1° segmento da seta /
                    LeaderAngle = Engine2.Geometry.AngleAdd(Angle, 45)
                    TempPoint = Engine2.Geometry.PolarPoint3d(StartPoint, LeaderAngle, Distance)
                    Points.Add(New Point3dCollection({StartPoint, TempPoint}))

                    'Adiciona o 2° segmento da seta |
                    TempAngle = Engine2.Geometry.AngleAdd(Angle, -90)
                    TempDist = (LeaderSize / 2)
                    TempPoint2 = Engine2.Geometry.PolarPoint3d(TempPoint, TempAngle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint, TempPoint2}))

                    'Adiciona o 3° segmento da seta _
                    TempDist = (Lenght - (LeaderSize * 2))
                    TempPoint = Engine2.Geometry.PolarPoint3d(TempPoint2, Angle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint2, TempPoint}))

                    'Adiciona o 4° segmento da seta |
                    TempAngle = Engine2.Geometry.AngleAdd(Angle, 90)
                    TempDist = (LeaderSize / 2)
                    TempPoint2 = Engine2.Geometry.PolarPoint3d(TempPoint, TempAngle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint, TempPoint2}))

                    'Adiciona o 5° segmento da seta \
                    Points.Add(New Point3dCollection({TempPoint2, EndPoint}))


                    '----------------------------------

                    'Reverte o ângulo
                    Angle = Engine2.Geometry.AngleAdd(Angle, 180)

                    '----------------------------------

                    'Adiciona o 1A° segmento da seta /
                    LeaderAngle = Engine2.Geometry.AngleAdd(Angle, 45)
                    TempPoint = Engine2.Geometry.PolarPoint3d(EndPoint, LeaderAngle, Distance)
                    Points.Add(New Point3dCollection({EndPoint, TempPoint}))

                    'Adiciona o 2A° segmento da seta |
                    TempAngle = Engine2.Geometry.AngleAdd(Angle, -90)
                    TempDist = (LeaderSize / 2)
                    TempPoint2 = Engine2.Geometry.PolarPoint3d(TempPoint, TempAngle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint, TempPoint2}))

                    'Adiciona o 3A° segmento da seta _
                    TempDist = (Lenght - (LeaderSize * 2))
                    TempPoint = Engine2.Geometry.PolarPoint3d(TempPoint2, Angle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint2, TempPoint}))

                    'Adiciona o 4A° segmento da seta |
                    TempAngle = Engine2.Geometry.AngleAdd(Angle, 90)
                    TempDist = (LeaderSize / 2)
                    TempPoint2 = Engine2.Geometry.PolarPoint3d(TempPoint, TempAngle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint, TempPoint2}))

                    'Adiciona o 5A° segmento da seta \
                    Points.Add(New Point3dCollection({TempPoint2, StartPoint}))

                Case eLeaderStyle.StartPointLeader

                    'Adiciona o 1° segmento da seta /
                    LeaderAngle = Engine2.Geometry.AngleAdd(Angle, 45)
                    TempPoint = Engine2.Geometry.PolarPoint3d(StartPoint, LeaderAngle, Distance)
                    Points.Add(New Point3dCollection({StartPoint, TempPoint}))

                    'Adiciona o 2° segmento da seta |
                    TempAngle = Engine2.Geometry.AngleAdd(Angle, -90)
                    TempDist = (LeaderSize / 2)
                    TempPoint2 = Engine2.Geometry.PolarPoint3d(TempPoint, TempAngle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint, TempPoint2}))

                    'Adiciona o 3° segmento da seta _
                    TempDist = (Lenght - LeaderSize)
                    TempPoint = Engine2.Geometry.PolarPoint3d(TempPoint2, Angle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint2, TempPoint}))

                    'Adiciona o 4° segmento da seta |
                    Points.Add(New Point3dCollection({TempPoint, EndPoint}))

                    '----------------------------------

                    'Reverte o ângulo
                    Angle = Engine2.Geometry.AngleAdd(Angle, 180)

                    '----------------------------------

                    'Adiciona o 1A° segmento da seta |
                    TempAngle = Engine2.Geometry.AngleAdd(Angle, 90)
                    TempDist = (LeaderSize / 2)
                    TempPoint2 = Engine2.Geometry.PolarPoint3d(EndPoint, TempAngle, TempDist)
                    Points.Add(New Point3dCollection({EndPoint, TempPoint2}))

                    'Adiciona o 2A° segmento da seta _
                    TempDist = (Lenght - LeaderSize)
                    TempPoint = Engine2.Geometry.PolarPoint3d(TempPoint2, Angle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint2, TempPoint}))

                    'Adiciona o 3A° segmento da seta |
                    TempAngle = Engine2.Geometry.AngleAdd(Angle, 90)
                    TempDist = (LeaderSize / 2)
                    TempPoint2 = Engine2.Geometry.PolarPoint3d(TempPoint, TempAngle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint, TempPoint2}))

                    'Adiciona o 4A° segmento da seta \
                    Points.Add(New Point3dCollection({TempPoint2, StartPoint}))

                Case eLeaderStyle.EndPointLeader

                    'Adiciona o 1° segmento da seta |
                    TempAngle = Engine2.Geometry.AngleAdd(Angle, -90)
                    TempDist = (LeaderSize / 2)
                    TempPoint2 = Engine2.Geometry.PolarPoint3d(StartPoint, TempAngle, TempDist)
                    Points.Add(New Point3dCollection({StartPoint, TempPoint2}))

                    'Adiciona o 2° segmento da seta _
                    TempDist = (Lenght - LeaderSize)
                    TempPoint = Engine2.Geometry.PolarPoint3d(TempPoint2, Angle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint2, TempPoint}))

                    'Adiciona o 3° segmento da seta |
                    TempAngle = Engine2.Geometry.AngleAdd(Angle, -90)
                    TempDist = (LeaderSize / 2)
                    TempPoint2 = Engine2.Geometry.PolarPoint3d(TempPoint, TempAngle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint, TempPoint2}))

                    'Adiciona o 4° segmento da seta \
                    Points.Add(New Point3dCollection({TempPoint2, EndPoint}))

                    '----------------------------------

                    'Reverte o ângulo
                    Angle = Engine2.Geometry.AngleAdd(Angle, 180)

                    '----------------------------------

                    'Adiciona o 1A° segmento da seta /
                    LeaderAngle = Engine2.Geometry.AngleAdd(Angle, -45)
                    TempPoint = Engine2.Geometry.PolarPoint3d(EndPoint, LeaderAngle, Distance)
                    Points.Add(New Point3dCollection({EndPoint, TempPoint}))

                    'Adiciona o 2A° segmento da seta |
                    TempAngle = Engine2.Geometry.AngleAdd(Angle, 90)
                    TempDist = (LeaderSize / 2)
                    TempPoint2 = Engine2.Geometry.PolarPoint3d(TempPoint, TempAngle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint, TempPoint2}))

                    'Adiciona o 3A° segmento da seta _
                    TempDist = (Lenght - LeaderSize)
                    TempPoint = Engine2.Geometry.PolarPoint3d(TempPoint2, Angle, TempDist)
                    Points.Add(New Point3dCollection({TempPoint2, TempPoint}))

                    'Adiciona o 4A° segmento da seta |
                    Points.Add(New Point3dCollection({TempPoint, StartPoint}))

            End Select

            'Desenha os vetores
            For Each Point3dCollection As Point3dCollection In Points
                Editor.DrawVector(Point3dCollection.Item(0), Point3dCollection.Item(1), ColorIndex, Highlight)
            Next

        End Sub

        ''' <summary>
        ''' Desenha o vetor entre dois pontos
        ''' </summary>
        ''' <param name="StartPoint">Ponto inicial</param>
        ''' <param name="EndPoint">Ponto final</param>
        ''' <param name="ColorIndex">Cor</param>
        ''' <param name="Highlight">Determina se o desenho será destacado</param>
        ''' <param name="Refresh">Determina se os vetores existentes serão excluídos durante o novo desenho</param>
        ''' <remarks></remarks>
        Public Shadows Sub DrawVector(StartPoint As Point3d, EndPoint As Point3d, ColorIndex As Integer, Optional Highlight As Boolean = False, Optional Refresh As Boolean = True)

            'Declarações
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            'Limpa os vetores anteriores
            If Refresh = True Then
                'Editor.Regen()
                Engine2.EntityInteration.Redraw2()
            End If

            'Desenha o vetor
            Editor.DrawVector(StartPoint, EndPoint, ColorIndex, Highlight)

        End Sub

        ''' <summary>
        ''' Desenha o vetor entre dois pontos
        ''' </summary>
        ''' <param name="Points">Coleção de pontos para o desenho dos vetores</param>
        ''' <param name="ColorIndex">Cor</param>
        ''' <param name="Highlight">Determina se o desenho será destacado</param>
        ''' <param name="Refresh">Determina se os vetores existentes serão excluídos durante o novo desenho</param>
        ''' <remarks></remarks>
        Public Shadows Sub DrawVector(Points As Point3dCollection, ColorIndex As Integer, Optional Highlight As Boolean = False, Optional Refresh As Boolean = True)

            'Declarações
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

            'Analisa a coleção
            If Points.Count = 1 Then
                Throw New System.Exception("O número de pontos informado não é válido.")
            End If

            'Limpa os vetores anteriores
            If Refresh = True Then
                'Editor.Regen()
                Engine2.EntityInteration.Redraw2()
            End If

            'Desenha os vetores
            For Index As Integer = 0 To Points.Count - 2
                Editor.DrawVector(Points.Item(Index), Points.Item(Index + 1), ColorIndex, Highlight)
            Next

        End Sub

        ''' <summary>
        ''' Desebha vetores com formato de caracter pré definido
        ''' </summary>
        ''' <param name="VectorChar">Formato</param>
        ''' <param name="Position">Posição</param>
        ''' <param name="FontHeight">Altura da fonte</param>
        ''' <param name="ColorIndex">Cor</param>
        ''' <param name="Highlight">Determina se o desenho será destacado</param>
        ''' <param name="Refresh">Determina se os vetores existentes serão excluídos durante o novo desenho</param>
        ''' <param name="UseUCS">Usar UCS</param>
        ''' <remarks></remarks>
        Public Shared Sub DrawCharVector(VectorChar As eVectorChar, Position As Point3d, FontHeight As Double, ColorIndex As Integer, Optional Highlight As Boolean = False, Optional Refresh As Boolean = True, Optional UseUCS As Boolean = True)

            'Declarações
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim Editor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            Dim Points0 As New Point3dCollection
            Dim Points1 As New Point3dCollection
            Dim Points2 As New Point3dCollection
            Dim Part As Double
            Dim Z As Double
            Dim P0 As Point3d
            Dim P1 As Point3d
            Dim P2 As Point3d
            Dim P3 As Point3d
            Dim P4 As Point3d
            Dim P5 As Point3d
            Dim P6 As Point3d
            Dim P7 As Point3d
            Dim P8 As Point3d
            Dim P9 As Point3d
            Dim P10 As Point3d
            Dim P11 As Point3d
            Dim P12 As Point3d
            Dim P13 As Point3d
            Dim P14 As Point3d
            Dim P15 As Point3d

            'Verifica se deve usar a UCS
            If UseUCS = True Then
                Position = Engine2.UCS.Point3dWcsToUcs(Position)
            End If

            'Obtem a dimensão da parte
            Part = (FontHeight / 6)

            'Seta o valor de Z
            Z = 0

            'Calcula os pontos
            With Points0
                .Clear()
                Select Case VectorChar
                    Case eVectorChar._0
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 2), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 4))
                        P2 = New Point3d((P1.X + Part), (P1.Y + Part), Z)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, (Part * 2))
                        P4 = New Point3d((P3.X + Part), (P3.Y - Part), Z)
                        P5 = Engine2.Geometry.PolarPoint3d(P4, 270, (Part * 4))
                        P6 = New Point3d((P5.X - Part), (P5.Y - Part), Z)
                        P7 = Engine2.Geometry.PolarPoint3d(P6, 180, (Part * 2))
                        P8 = P0
                        P9 = New Point3d(P8.X + (Part * 4), P8.Y + (Part * 4), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                        .Add(P8)
                        .Add(P9)
                    Case eVectorChar._1
                        P0 = New Point3d(Position.X + (Part / 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 6))
                        P2 = New Point3d(P1.X - Part, P1.Y - Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                    Case eVectorChar._2
                        P0 = New Point3d(Position.X + (Part * 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 180, (Part * 4))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 90, Part)
                        P3 = New Point3d(P2.X + (Part * 4), P2.Y + (Part * 3), Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 90, Part)
                        P5 = New Point3d((P4.X - Part), (P4.Y + Part), Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 180, (Part * 2))
                        P7 = New Point3d((P6.X - Part), (P6.Y - Part), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                    Case eVectorChar._3
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 2), Z)
                        P1 = New Point3d(P0.X + Part, P0.Y + Part, Z)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 0, (Part * 2))
                        P3 = New Point3d(P2.X + Part, P2.Y - Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 270, Part)
                        P5 = New Point3d(P4.X - Part, P4.Y - Part, Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 180, (Part * 2))
                        P7 = Engine2.Geometry.PolarPoint3d(P6, 0, (Part * 2))
                        P8 = New Point3d(P7.X + Part, P7.Y - Part, Z)
                        P9 = Engine2.Geometry.PolarPoint3d(P8, 270, Part)
                        P10 = New Point3d(P9.X - Part, P9.Y - Part, Z)
                        P11 = Engine2.Geometry.PolarPoint3d(P10, 180, (Part * 2))
                        P12 = New Point3d(P11.X - Part, P11.Y + Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                        .Add(P8)
                        .Add(P9)
                        .Add(P10)
                        .Add(P11)
                        .Add(P12)
                    Case eVectorChar._4
                        P0 = New Point3d(Position.X + (Part * 2), Position.Y - Part, Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 180, (Part * 4))
                        P2 = New Point3d(P1.X + (Part * 2), P1.Y + (Part * 4), Z)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, Part)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 270, (Part * 6))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                    Case eVectorChar._5
                        P0 = New Point3d(Position.X + (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 180, (Part * 4))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, (Part * 3))
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, (Part * 3))
                        P4 = New Point3d(P3.X + Part, P3.Y - Part, Z)
                        P5 = Engine2.Geometry.PolarPoint3d(P4, 270, Part)
                        P6 = New Point3d(P5.X - Part, P5.Y - Part, Z)
                        P7 = Engine2.Geometry.PolarPoint3d(P6, 180, (Part * 3))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                    Case eVectorChar._6
                        P0 = New Point3d(Position.X + (Part * 2), Position.Y + (Part * 2), Z)
                        P1 = New Point3d(P0.X - Part, P0.Y + Part, Z)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 180, (Part * 2))
                        P3 = New Point3d(P2.X - Part, P2.Y - Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 270, (Part * 4))
                        P5 = New Point3d(P4.X + Part, P4.Y - Part, Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 0, (Part * 2))
                        P7 = New Point3d(P6.X + Part, P6.Y + Part, Z)
                        P8 = Engine2.Geometry.PolarPoint3d(P7, 90, Part)
                        P9 = New Point3d(P8.X - Part, P8.Y + Part, Z)
                        P10 = Engine2.Geometry.PolarPoint3d(P9, 180, (Part * 2))
                        P11 = New Point3d(P10.X - Part, P10.Y - Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                        .Add(P8)
                        .Add(P9)
                        .Add(P10)
                        .Add(P11)
                    Case eVectorChar._7
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, (Part * 4))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, Part)
                        P3 = New Point3d(P2.X - (Part * 4), P2.Y - (Part * 5), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                    Case eVectorChar._8
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 2), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, Part)
                        P2 = New Point3d(P1.X + Part, P1.Y + Part, Z)
                        P3 = New Point3d(P2.X - Part, P2.Y + Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 90, Part)
                        P5 = New Point3d(P4.X + Part, P4.Y + Part, Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 0, (Part * 2))
                        P7 = New Point3d(P6.X + Part, P6.Y - Part, Z)
                        P8 = Engine2.Geometry.PolarPoint3d(P7, 270, Part)
                        P9 = New Point3d(P8.X - Part, P8.Y - Part, Z)
                        P10 = Engine2.Geometry.PolarPoint3d(P9, 180, (Part * 2))
                        P11 = Engine2.Geometry.PolarPoint3d(P10, 0, (Part * 2))
                        P12 = New Point3d(P11.X + Part, P11.Y - Part, Z)
                        P13 = Engine2.Geometry.PolarPoint3d(P12, 270, Part)
                        P14 = New Point3d(P13.X - Part, P13.Y - Part, Z)
                        P15 = Engine2.Geometry.PolarPoint3d(P14, 180, (Part * 2))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                        .Add(P8)
                        .Add(P9)
                        .Add(P10)
                        .Add(P11)
                        .Add(P12)
                        .Add(P13)
                        .Add(P14)
                        .Add(P15)
                        .Add(P0)
                    Case eVectorChar._9
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 2), Z)
                        P1 = New Point3d(P0.X + Part, P0.Y - Part, Z)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 0, (Part * 2))
                        P3 = New Point3d(P2.X + Part, P2.Y + Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 90, (Part * 4))
                        P5 = New Point3d(P4.X - Part, P4.Y + Part, Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 180, (Part * 2))
                        P7 = New Point3d(P6.X - Part, P6.Y - Part, Z)
                        P8 = Engine2.Geometry.PolarPoint3d(P7, 270, Part)
                        P9 = New Point3d(P8.X + Part, P8.Y - Part, Z)
                        P10 = Engine2.Geometry.PolarPoint3d(P9, 0, (Part * 2))
                        P11 = New Point3d(P10.X + Part, P10.Y + Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                        .Add(P8)
                        .Add(P9)
                        .Add(P10)
                        .Add(P11)
                    Case eVectorChar.PONTO
                        P0 = New Point3d(Position.X - (Part / 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, Part)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 0, Part)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 270, Part)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P0)
                    Case eVectorChar.VIRGULA
                        P0 = New Point3d(Position.X + (Part / 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 180, Part)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 90, Part)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, Part)
                        P4 = P0
                        P5 = New Point3d(P4.X - Part, P4.Y - Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                    Case eVectorChar.ABRE_PARENTESE
                        P0 = New Point3d(Position.X + (Part / 2), Position.Y + (Part * 3), Z)
                        P1 = New Point3d(P0.X - Part, P0.Y - Part, Z)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, (Part * 4))
                        P3 = New Point3d(P2.X + Part, P2.Y - Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                    Case eVectorChar.FECHA_PARENTESE
                        P0 = New Point3d(Position.X - (Part / 2), Position.Y + (Part * 3), Z)
                        P1 = New Point3d(P0.X + Part, P0.Y - Part, Z)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, (Part * 4))
                        P3 = New Point3d(P2.X - Part, P2.Y - Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                    Case eVectorChar.ABRE_COLCHETE
                        P0 = New Point3d(Position.X + (Part / 2), Position.Y + (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 180, Part)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, (Part * 6))
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, Part)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                    Case eVectorChar.FECHA_COLCHETE
                        P0 = New Point3d(Position.X - (Part / 2), Position.Y + (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, Part)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, (Part * 6))
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 180, Part)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                    Case eVectorChar.ABRE_CHAVE
                        P0 = New Point3d(Position.X + Part, Position.Y + (Part * 3), Z)
                        P1 = New Point3d(P0.X - Part, P0.Y - Part, Z)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, Part)
                        P3 = New Point3d(P2.X - Part, P2.Y - Part, Z)
                        P4 = New Point3d(P3.X + Part, P3.Y - Part, Z)
                        P5 = Engine2.Geometry.PolarPoint3d(P4, 270, Part)
                        P6 = New Point3d(P5.X + Part, P5.Y - Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                    Case eVectorChar.FECHA_CHAVE
                        P0 = New Point3d(Position.X - Part, Position.Y + (Part * 3), Z)
                        P1 = New Point3d(P0.X + Part, P0.Y - Part, Z)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, Part)
                        P3 = New Point3d(P2.X + Part, P2.Y - Part, Z)
                        P4 = New Point3d(P3.X - Part, P3.Y - Part, Z)
                        P5 = Engine2.Geometry.PolarPoint3d(P4, 270, Part)
                        P6 = New Point3d(P5.X - Part, P5.Y - Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                    Case eVectorChar.SOMA
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y, Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, (Part * 4))
                        P2 = Position
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 90, (Part * 2))
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 270, (Part * 4))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                    Case eVectorChar.DIVISAO
                        P0 = New Point3d(Position.X - (Part * 1.5), Position.Y, Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, (Part * 3))
                        .Add(P0)
                        .Add(P1)
                        P0 = New Point3d(Position.X - (Part / 2), Position.Y + Part, Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, Part)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 90, Part)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 180, Part)
                        P4 = P0
                        Points1.Add(P0)
                        Points1.Add(P1)
                        Points1.Add(P2)
                        Points1.Add(P3)
                        Points1.Add(P4)
                        P0 = New Point3d(Position.X - (Part / 2), Position.Y - Part, Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, Part)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, Part)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 180, Part)
                        P4 = P0
                        Points2.Add(P0)
                        Points2.Add(P1)
                        Points2.Add(P2)
                        Points2.Add(P3)
                        Points2.Add(P4)
                    Case eVectorChar.MULTIPLICACAO
                        P0 = New Point3d(Position.X - (Part * 1.5), Position.Y + (Part * 1.5), Z)
                        P1 = New Point3d(P0.X + (Part * 3), P0.Y - (Part * 3), Z)
                        P2 = New Point3d(P1.X - (Part * 1.5), P1.Y + (Part * 1.5), Z)
                        P3 = New Point3d(P2.X - (Part * 1.5), P2.Y - (Part * 1.5), Z)
                        P4 = New Point3d(P3.X + (Part * 3), P3.Y + (Part * 3), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                    Case eVectorChar.MAIOR
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = New Point3d(P0.X + (Part * 4), P0.Y - (Part * 3), Z)
                        P2 = New Point3d(P1.X - (Part * 4), P1.Y - (Part * 3), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                    Case eVectorChar.MENOR
                        P0 = New Point3d(Position.X + (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = New Point3d(P0.X - (Part * 4), P0.Y - (Part * 3), Z)
                        P2 = New Point3d(P1.X + (Part * 4), P1.Y - (Part * 3), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                    Case eVectorChar.IGUAL
                        P0 = New Point3d(Position.X - (Part * 1.5), Position.Y + (Part / 2), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, (Part * 3))
                        .Add(P0)
                        .Add(P1)
                        P0 = New Point3d(Position.X - (Part * 1.5), Position.Y - (Part / 2), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, (Part * 3))
                        Points1.Add(P0)
                        Points1.Add(P1)
                    Case eVectorChar.SUBTRACAO
                        P0 = New Point3d(Position.X - (Part * 1.5), Position.Y, Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, (Part * 3))
                        .Add(P0)
                        .Add(P1)
                    Case eVectorChar.DIFERENTE
                        P0 = New Point3d(Position.X - (Part * 1.5), Position.Y + (Part / 2), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, (Part * 3))
                        .Add(P0)
                        .Add(P1)
                        P0 = New Point3d(Position.X - (Part * 1.5), Position.Y - (Part / 2), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, (Part * 3))
                        Points1.Add(P0)
                        Points1.Add(P1)
                        P0 = New Point3d(Position.X - (Part * 1.5), Position.Y + (Part * 2), Z)
                        P1 = New Point3d(P0.X + (Part * 3), P0.Y - (Part * 4), Z)
                        Points2.Add(P0)
                        Points2.Add(P1)
                    Case eVectorChar.PORCENTAGEM
                        P0 = New Point3d(Position.X + (Part * 1.5), Position.Y + (Part * 1.5), Z)
                        P1 = New Point3d(P0.X - (Part * 3), P0.Y - (Part * 3), Z)
                        .Add(P0)
                        .Add(P1)
                        P0 = New Point3d(Position.X - (Part / 2), Position.Y + (Part / 2), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 180, Part)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 90, Part)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, Part)
                        P4 = P0
                        Points1.Add(P0)
                        Points1.Add(P1)
                        Points1.Add(P2)
                        Points1.Add(P3)
                        Points1.Add(P4)
                        P0 = New Point3d(Position.X + (Part / 2), Position.Y - (Part / 2), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, Part)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, Part)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 180, Part)
                        P4 = P0
                        Points2.Add(P0)
                        Points2.Add(P1)
                        Points2.Add(P2)
                        Points2.Add(P3)
                        Points2.Add(P4)
                    Case eVectorChar.A
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 3), Z)
                        P1 = New Point3d(P0.X + (Part * 2), P0.Y + (Part * 6), Z)
                        P2 = New Point3d(P1.X + Part, P1.Y - (Part * 3), Z)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 180, (Part * 2))
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 0, (Part * 2))
                        P5 = New Point3d(P4.X + Part, P4.Y - (Part * 3), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                    Case eVectorChar.B
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 6))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 0, (Part * 3))
                        P3 = New Point3d(P2.X + Part, P2.Y - Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 270, Part)
                        P5 = New Point3d(P4.X - Part, P4.Y - Part, Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 180, (Part * 3))
                        P7 = Engine2.Geometry.PolarPoint3d(P6, 0, (Part * 3))
                        P8 = New Point3d(P7.X + Part, P7.Y - Part, Z)
                        P9 = Engine2.Geometry.PolarPoint3d(P8, 270, Part)
                        P10 = New Point3d(P9.X - Part, P9.Y - Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                        .Add(P8)
                        .Add(P9)
                        .Add(P10)
                        .Add(P0)
                    Case eVectorChar.C
                        P0 = New Point3d(Position.X + (Part * 2), Position.Y + (Part * 2), Z)
                        P1 = New Point3d(P0.X - Part, P0.Y + Part, Z)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 180, (Part * 2))
                        P3 = New Point3d(P2.X - Part, P2.Y - Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 270, (Part * 4))
                        P5 = New Point3d(P4.X + Part, P4.Y - Part, Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 0, (Part * 2))
                        P7 = New Point3d(P6.X + Part, P6.Y + Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                    Case eVectorChar.D
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 6))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 0, (Part * 3))
                        P3 = New Point3d(P2.X + Part, P2.Y - Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 270, (Part * 4))
                        P5 = New Point3d(P4.X - Part, P4.Y - Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P0)
                    Case eVectorChar.E
                        P0 = New Point3d(Position.X + (Part * 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 180, (Part * 4))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 90, (Part * 3))
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, (Part * 3))
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 180, (Part * 3))
                        P5 = Engine2.Geometry.PolarPoint3d(P4, 90, (Part * 3))
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 0, (Part * 4))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                    Case eVectorChar.F
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 3))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 0, (Part * 3))
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 180, (Part * 3))
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 90, (Part * 3))
                        P5 = Engine2.Geometry.PolarPoint3d(P4, 0, (Part * 4))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                    Case eVectorChar.G
                        P0 = New Point3d(Position.X + (Part * 2), Position.Y + (Part * 2), Z)
                        P1 = New Point3d(P0.X - Part, P0.Y + Part, Z)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 180, (Part * 2))
                        P3 = New Point3d(P2.X - Part, P2.Y - Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 270, (Part * 4))
                        P5 = New Point3d(P4.X + Part, P4.Y - Part, Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 0, (Part * 2))
                        P7 = New Point3d(P6.X + Part, P6.Y + Part, Z)
                        P8 = Engine2.Geometry.PolarPoint3d(P7, 90, (Part * 2))
                        P9 = Engine2.Geometry.PolarPoint3d(P8, 180, (Part * 3))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                        .Add(P8)
                        .Add(P9)
                    Case eVectorChar.H
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 6))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, (Part * 3))
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, (Part * 4))
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 90, (Part * 3))
                        P5 = Engine2.Geometry.PolarPoint3d(P4, 270, (Part * 6))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                    Case eVectorChar.I
                        P0 = New Point3d(Position.X, Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 6))
                        .Add(P0)
                        .Add(P1)
                    Case eVectorChar.J
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, (Part * 4))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, (Part * 5))
                        P3 = New Point3d(P2.X - Part, P2.Y - Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 180, (Part * 2))
                        P5 = New Point3d(P4.X - Part, P4.Y + Part, Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 90, (Part * 2))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                    Case eVectorChar.K
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 270, (Part * 3))
                        P2 = New Point3d(P1.X + (Part * 4), P1.Y + (Part * 3), Z)
                        P3 = P1
                        P4 = New Point3d(P3.X + (Part * 4), P3.Y - (Part * 3), Z)
                        P5 = P1
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 270, (Part * 3))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                    Case eVectorChar.L
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 270, (Part * 6))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 0, (Part * 4))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                    Case eVectorChar.M
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 6))
                        P2 = New Point3d(P1.X + (Part * 2), P1.Y - (Part * 6), Z)
                        P3 = New Point3d(P2.X + (Part * 2), P2.Y + (Part * 6), Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 270, (Part * 6))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                    Case eVectorChar.N
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 6))
                        P2 = New Point3d(P1.X + (Part * 4), P1.Y - (Part * 6), Z)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 90, (Part * 6))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                    Case eVectorChar.O
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 2), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 4))
                        P2 = New Point3d((P1.X + Part), (P1.Y + Part), Z)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, (Part * 2))
                        P4 = New Point3d((P3.X + Part), (P3.Y - Part), Z)
                        P5 = Engine2.Geometry.PolarPoint3d(P4, 270, (Part * 4))
                        P6 = New Point3d((P5.X - Part), (P5.Y - Part), Z)
                        P7 = Engine2.Geometry.PolarPoint3d(P6, 180, (Part * 2))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                        .Add(P0)
                    Case eVectorChar.P
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 6))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 0, (Part * 3))
                        P3 = New Point3d(P2.X + Part, P2.Y - Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 270, Part)
                        P5 = New Point3d(P4.X - Part, P4.Y - Part, Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 180, (Part * 3))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                    Case eVectorChar.Q
                        P0 = New Point3d(Position.X + Part, Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 180, (Part * 2))
                        P2 = New Point3d(P1.X - Part, P1.Y + Part, Z)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 90, (Part * 4))
                        P4 = New Point3d(P3.X + Part, P3.Y + Part, Z)
                        P5 = Engine2.Geometry.PolarPoint3d(P4, 0, (Part * 2))
                        P6 = New Point3d(P5.X + Part, P5.Y - Part, Z)
                        P7 = Engine2.Geometry.PolarPoint3d(P6, 270, (Part * 4))
                        P8 = New Point3d(P7.X - Part, P7.Y - Part, Z)
                        P9 = New Point3d(P8.X - Part, P8.Y + Part, Z)
                        P10 = New Point3d(P9.X + (Part * 2), P9.Y - (Part * 2), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                        .Add(P8)
                        .Add(P9)
                        .Add(P10)
                    Case eVectorChar.R
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 6))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 0, (Part * 3))
                        P3 = New Point3d(P2.X + Part, P2.Y - Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 270, Part)
                        P5 = New Point3d(P4.X - Part, P4.Y - Part, Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 180, (Part * 3))
                        P7 = Engine2.Geometry.PolarPoint3d(P6, 0, (Part * 3))
                        P8 = New Point3d(P7.X + Part, P7.Y - Part, Z)
                        P9 = Engine2.Geometry.PolarPoint3d(P8, 270, (Part * 2))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                        .Add(P8)
                        .Add(P9)
                    Case eVectorChar.S
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y - (Part * 2), Z)
                        P1 = New Point3d(P0.X + Part, P0.Y - Part, Z)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 0, (Part * 2))
                        P3 = New Point3d(P2.X + Part, P2.Y + Part, Z)
                        P4 = Engine2.Geometry.PolarPoint3d(P3, 90, Part)
                        P5 = New Point3d(P4.X - Part, P4.Y + Part, Z)
                        P6 = Engine2.Geometry.PolarPoint3d(P5, 180, (Part * 2))
                        P7 = New Point3d(P6.X - Part, P6.Y + Part, Z)
                        P8 = Engine2.Geometry.PolarPoint3d(P7, 90, Part)
                        P9 = New Point3d(P8.X + Part, P8.Y + Part, Z)
                        P10 = Engine2.Geometry.PolarPoint3d(P9, 0, (Part * 2))
                        P11 = New Point3d(P10.X + Part, P10.Y - Part, Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                        .Add(P6)
                        .Add(P7)
                        .Add(P8)
                        .Add(P9)
                        .Add(P10)
                        .Add(P11)
                    Case eVectorChar.T
                        P0 = New Point3d(Position.X, Position.Y - (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 90, (Part * 6))
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 180, (Part * 2))
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, (Part * 4))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                    Case eVectorChar.U
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 270, (Part * 5))
                        P2 = New Point3d(P1.X + Part, P1.Y - Part, Z)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, (Part * 2))
                        P4 = New Point3d(P3.X + Part, P3.Y + Part, Z)
                        P5 = Engine2.Geometry.PolarPoint3d(P4, 90, (Part * 5))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                        .Add(P5)
                    Case eVectorChar.V
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = New Point3d(P0.X + (Part * 2), P0.Y - (Part * 6), Z)
                        P2 = New Point3d(P1.X + (Part * 2), P1.Y + (Part * 6), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                    Case eVectorChar.X
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = New Point3d(P0.X + (Part * 4), P0.Y - (Part * 6), Z)
                        P2 = New Point3d(P1.X - (Part * 2), P1.Y + (Part * 3), Z)
                        P3 = New Point3d(P2.X - (Part * 2), P2.Y - (Part * 3), Z)
                        P4 = New Point3d(P3.X + (Part * 4), P3.Y + (Part * 6), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                    Case eVectorChar.Y
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = New Point3d(P0.X + (Part * 2), P0.Y - (Part * 3), Z)
                        P2 = Engine2.Geometry.PolarPoint3d(P1, 270, (Part * 3))
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 90, (Part * 3))
                        P4 = New Point3d(P3.X + (Part * 2), P3.Y + (Part * 3), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                    Case eVectorChar.W
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = New Point3d(P0.X + Part, P0.Y - (Part * 6), Z)
                        P2 = New Point3d(P1.X + Part, P1.Y + (Part * 6), Z)
                        P3 = New Point3d(P2.X + Part, P2.Y - (Part * 6), Z)
                        P4 = New Point3d(P3.X + Part, P3.Y + (Part * 6), Z)
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                        .Add(P4)
                    Case eVectorChar.Z
                        P0 = New Point3d(Position.X - (Part * 2), Position.Y + (Part * 3), Z)
                        P1 = Engine2.Geometry.PolarPoint3d(P0, 0, (Part * 4))
                        P2 = New Point3d(P1.X - (Part * 4), P1.Y - (Part * 6), Z)
                        P3 = Engine2.Geometry.PolarPoint3d(P2, 0, (Part * 4))
                        .Add(P0)
                        .Add(P1)
                        .Add(P2)
                        .Add(P3)
                End Select
            End With

            'Limpa os vetores anteriores
            If Refresh = True Then
                'Editor.Regen()
                Engine2.EntityInteration.Redraw2()
            End If

            'Desenha os vetores
            For Index As Integer = 0 To Points0.Count - 2
                Editor.DrawVector(Points0.Item(Index), Points0.Item(Index + 1), ColorIndex, Highlight)
            Next
            For Index As Integer = 0 To Points1.Count - 2
                Editor.DrawVector(Points1.Item(Index), Points1.Item(Index + 1), ColorIndex, Highlight)
            Next
            For Index As Integer = 0 To Points2.Count - 2
                Editor.DrawVector(Points2.Item(Index), Points2.Item(Index + 1), ColorIndex, Highlight)
            Next

        End Sub

#End Region



    End Class

End Namespace

